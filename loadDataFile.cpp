#include "loadDataFile.h"
#include <windows.h>
#include <QDebug>

LoadDataFile::LoadDataFile(QObject *parent) : QObject(parent)
{
	_mExcelReader = new ExcelReader();
	_filePath = "";
	_dataOperateType = "";
	_path = QCoreApplication::applicationDirPath().toUtf8();
	_uid = "";
	_info = new MEvaluationInfo();
}

LoadDataFile::~LoadDataFile()
{

}

void LoadDataFile::readDataFile(QList<MResult*>* results, QString filePath)
{
	_results = results;
	_filePath = filePath;
	_dataOperateType = "readResultData";
}

void LoadDataFile::setInfoData(MEvaluationInfo* evaluationInfo)
{
	_info = evaluationInfo;
}

void LoadDataFile::setExcelData(MResult* result, QString excelName)
{
	_result = result;
	_excelName = excelName;
	_dataOperateType = "setExcelData";
}

void LoadDataFile::getCount(int rec, int sum)
{
    emit countStep(rec, sum);
}

void LoadDataFile::doDataOperate()
{
    if(_dataOperateType == "readResultData")
    {
		readResultsFromFile();

        emit finish();
    }
	else if (_dataOperateType == "setExcelData")
	{
		generateExcelResult(_results, _excelName);

		emit finishExcel();
	}
    else
    {
		
    }
}

void LoadDataFile::readResultsFromFile()
{
	QString appPath = QCoreApplication::applicationDirPath().toStdString().data();
	QFile file(_filePath);
	if (file.exists())
	{
		file.open(QIODevice::ReadOnly);
		QDataStream input(&file);
		QString uid;
		int lenResults;
		input >> _uid >> lenResults;
		for (int patternCount = 0; patternCount < lenResults; patternCount++)
		{
			MResult* rst = new MResult;
			_results->append(rst);
			
			//MResult* rst = _results->at(patternCount);
			input >> rst->PatternGUID >> rst->_isInitialized >> rst->_validFormCount >> rst->FormResultsLen;
			int len = rst->FormResultsLen;
			for (int i = 0; i < len; i++)
			{
				getCount(i, len);
				MFormResult* mfrst = new MFormResult();
				_results->at(patternCount)->AddFormResult(mfrst);
				QByteArray mfrstPic;
				input >> mfrst->DeviceIndex >> mfrst->QueueIndex >> mfrst->IsBlankPaper >> mfrst->IsRecognizeSuccess >> mfrst->IsForwardDirection \
					>> mfrst->FormIndex >> mfrst->IsRectified >> mfrst->LenMarkGroupResults >> mfrst->LenImageShotResults >> mfrst->ErrorReason >> mfrst->ImgUrl >> mfrstPic;
				//QPixmap img1(mfrst->ImgUrl);
				QImage img;
				img.loadFromData(mfrstPic);
				QString imgUrl = _path + "/imageGray/" + mfrst->ImgUrl;
				// 避免重名
				QString changedImageFileName;
				QString fileName = imgUrl.left(imgUrl.lastIndexOf("."));
				std::fstream tempFile;
				//ELOGI("start detect  multiple name   " << changedImageFileName.toStdString());
				int imageFileIndex = 0;
				while (true)
				{
					tempFile.open(imgUrl.toLocal8Bit(), std::ios::in);
					if (tempFile)
					{
						//ELOGI("find multiple name " );
						changedImageFileName = fileName + "(" + QString::number(imageFileIndex) + ").jpg";
						imageFileIndex++;
						//ELOGI("find multiple name   " << changedImageFileName.toStdString());
						tempFile.close();
						imgUrl = changedImageFileName;
					}
					else
					{
						//ELOGI("didn't find multiple name ");
						break;
					}
				}
				img.save(imgUrl);
				//查找imgurl中从后往前的第一个/位置
				int first = imgUrl.lastIndexOf("/"); //从后面查找"/"位置
				QString title = imgUrl.right(imgUrl.length() - first - 1); //从右边截取
				mfrst->ImgUrl = title;
				int lenMarks = mfrst->LenMarkGroupResults;
				for (int j = 0; j < lenMarks; j++)
				{
					MGroupResult* mgrt = new MGroupResult();
					input >> mgrt->IsValid >> mgrt->GroupName >> mgrt->TextResult >> mgrt->NumberResult >> mgrt->LenCellResults;
					int lenCells = mgrt->LenCellResults;
					for (int k = 0; k < lenCells; k++)
					{
						MCellResult* mcrst = new MCellResult();
						int kk;
						input >> mcrst->CellIndex >> mcrst->RowIndex >> mcrst->ColIndex >> mcrst->CellName \
							>> mcrst->CellRect.X >> mcrst->CellRect.Y >> mcrst->CellRect.Width >> mcrst->CellRect.Height >> kk;
						mcrst->FillSymbolType = Cm3::FormResult::MarkSymbolResult(kk);
						mgrt->AddCellResult(mcrst);
					}
					mfrst->AddGroupResult(mgrt);
				}
				for (int j = 0; j < mfrst->LenImageShotResults; j++)
				{
					MShotResult* msrst = new MShotResult();
					input >> msrst->GroupName >> msrst->ShotImageByteArray >> msrst->CellRect.X >> msrst->CellRect.Y >> msrst->CellRect.Width >> msrst->CellRect.Height;
					mfrst->AddShotResult(msrst);
				}
				//rst->AddFormResultFromBinaryFile(mfrst);
				//读出标志码识别结果
				MRegionResult* mrgrst = new MRegionResult();
				input >> mrgrst->Name >> mrgrst->X >> mrgrst->Y >> mrgrst->Width >> mrgrst->Height >> mrgrst->Result;
				mfrst->IdentifierResult = mrgrst;
			}
		}
	}
	file.close();
	//MFormResult::saveToFile(Path,*this->FormReuslts);
}

void LoadDataFile::generateExcelResult(QList<MResult*>* results, QString excelName)
{
	//规定excel的单元格式样式
	QXlsx::Format format1;
	format1.setFontColor(QColor(Qt::black));
	format1.setPatternBackgroundColor(QColor(Qt::white));
	format1.setFontSize(15);
	format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
	format1.setBorderStyle(QXlsx::Format::BorderMedium);/*边框样式*/
	MResult* currentResult;
	//直接进行保存
	QString evaluationName = _info->RecognizePatternInfo->Name;
	QString excelFileName = evaluationName + ".xlsx";
	//如果文件名重复
	QDir* dir = new QDir(excelName);
	QStringList filter;
	QList<QFileInfo>* fileInfo = new QList<QFileInfo>(dir->entryInfoList(filter));
	int repeatCount = 0, fileCount = 0;
	bool repeatFlag = 1;
	while (repeatFlag)
	{
		for (fileCount = 0; fileCount < fileInfo->count(); fileCount++)
		{
			if (excelFileName == fileInfo->at(fileCount).fileName()) {
				repeatCount++;
				excelFileName = evaluationName + "(" + QString::number(repeatCount) + ").xlsx";
				break;
			}
		}
		if (fileCount == fileInfo->count())
		{
			repeatFlag = 0;
		}
	}
	QString excelTempName = excelName + "/" + excelFileName;
	if (_mExcelReader->newExcel(excelTempName))
	{
		_mExcelReader->setSheetName(evaluationName);

		//初始化表头
		//单位页
		_mExcelReader->chooseSheet(0);
		_mExcelReader->writeExcel(1, 1, u8"单位指标汇总表", format1);

		_mExcelReader->writeExcel(2, 1, u8"单位", format1);
		_mExcelReader->writeExcel(2, 2, u8"表类型", format1);
		_mExcelReader->writeExcel(2, 3, u8"收回数", format1);

		_mExcelReader->mergeCells(2, 1, 4, 1, format1);
		_mExcelReader->mergeCells(2, 2, 4, 2, format1);
		_mExcelReader->mergeCells(2, 3, 4, 3, format1);
		//遍历模式
		int patternCount = _info->RecognizePatternInfo->RecognizeFormPatterns->count();
		QList<QString> patternName;
		patternName.clear();
		for (unsigned i = 0; i < patternCount; i++)
		{
			patternName.append(_info->RecognizePatternInfo->RecognizeFormPatterns->at(i)->FileName);
			_mExcelReader->writeExcel(i + 5, 2, patternName.at(i), format1);
		}
		//遍历单位
		_mExcelReader->writeExcel(patternCount + 5, 2, u8"合计", format1);
		int unitCount = _info->RecognizePatternInfo->RecognizeFormPatterns->at(0)->EvaluationUnits.count();
		QList<QString> unitName;
		unitName.clear();
		for (unsigned i = 0; i < unitCount; i++)
		{
			unitName.append(_info->RecognizePatternInfo->RecognizeFormPatterns->at(0)->EvaluationUnits.at(i));
			_mExcelReader->writeExcel(i * (patternCount + 1) + 5, 1, unitName.at(i), format1);
			_mExcelReader->mergeCells(i * (patternCount + 1) + 5, 1, (i + 1)*(patternCount + 1) + 4, 1, format1);

		}
		//遍历指标
		int indexCountF = _info->RecognizePatternInfo->RecognizeFormPatterns->at(0)->MemberIndexs->at(0)->MemberDetailIndexs->count();
		int indexCountS = 4;
		QList<QString> indexNameF;
		QList<MemberDetailIndex*>* indexNode = new QList<MemberDetailIndex*>;
		indexNameF.clear();
		indexNode->clear();
		for (unsigned i = 0; i < indexCountF; i ++)
		{
			indexNode->append(_info->RecognizePatternInfo->RecognizeFormPatterns->at(0)->MemberIndexs->at(0)->MemberDetailIndexs->at(i));
			//_info->RecognizePatternInfo->RecognizeFormPatterns->at(0)->MemberIndexs->at(0)->MemberDetailIndexs->at()->FirstLevelIndex.name
			indexNameF.append(indexNode->at(i)->FirstLevelIndex.name);
			_mExcelReader->writeExcel(2, indexCountS, indexNameF.at(i), format1);
			//二级指标
			unsigned j;
			for (j = 0; j < indexNode->at(i)->SecondLevelIndex.count(); j++)
			{
				_mExcelReader->writeExcel(4, indexCountS, indexNode->at(i)->SecondLevelIndex.at(j).name, format1);
				indexCountS++;
			}
			if (j == 0)
			{
				_mExcelReader->mergeCells(2, indexCountS, 4, indexCountS, format1);
				indexCountS ++;
			}
			else
			{
				_mExcelReader->mergeCells(2, indexCountS - j, 3, indexCountS - 1, format1);
			}
		}

		//_mExcelReader->mergeCells(1, 1, 1, 3, format1);//无指标情况
		_mExcelReader->mergeCells(1, 1, 1, indexCountS - 1, format1);
		//计票人


		_mExcelReader->chooseSheet(1);
		_mExcelReader->writeExcel(1, 1, u8"成员指标汇总表", format1);
		_mExcelReader->writeExcel(2, 1, u8"单位", format1);
		_mExcelReader->writeExcel(2, 2, u8"姓名", format1);
		_mExcelReader->writeExcel(2, 3, u8"性别", format1);
		_mExcelReader->writeExcel(2, 4, u8"职务", format1);
		_mExcelReader->writeExcel(2, 5, u8"职务类型", format1);
		_mExcelReader->writeExcel(2, 6, u8"表类型", format1);
		_mExcelReader->writeExcel(2, 7, u8"收回数", format1);
		//index
		indexCountS = 8;
		for (unsigned i = 0; i < indexCountF; i++)
		{
			_mExcelReader->writeExcel(2, indexCountS, indexNameF.at(i), format1);
			//二级指标
			unsigned j;
			for (j = 0; j < indexNode->at(i)->SecondLevelIndex.count(); j++)
			{
				_mExcelReader->writeExcel(4, indexCountS, indexNode->at(i)->SecondLevelIndex.at(j).name, format1);
				indexCountS++;
			}
			if (j == 0)
			{
				_mExcelReader->mergeCells(2, indexCountS, 4, indexCountS, format1);
				indexCountS++;
			}
			else
			{
				_mExcelReader->mergeCells(2, indexCountS - j, 3, indexCountS - 1, format1);
			}
		}

		_mExcelReader->mergeCells(1, 1, 1, indexCountS - 1, format1);
		_mExcelReader->mergeCells(2, 1, 4, 1, format1);
		_mExcelReader->mergeCells(2, 2, 4, 2, format1);
		_mExcelReader->mergeCells(2, 3, 4, 3, format1);
		_mExcelReader->mergeCells(2, 4, 4, 4, format1);
		_mExcelReader->mergeCells(2, 5, 4, 5, format1);
		_mExcelReader->mergeCells(2, 6, 4, 6, format1);
		_mExcelReader->mergeCells(2, 7, 4, 7, format1);
	
		for (int i = 0, takenSpace = 5 ; i < unitCount; i++)
		{
			int memberCount = _info->EvaluationMemberInfo->at(0)->EvaluationMembers->at(i)->EvaluationMembers->count();
			_mExcelReader->writeExcel(takenSpace, 1, unitName.at(i), format1);
			for (int j = 0; j < memberCount; j++)
			{
				MemberInfo* newMember = _info->EvaluationMemberInfo->at(0)->EvaluationMembers->at(i)->EvaluationMembers->at(j);
				_mExcelReader->writeExcel(takenSpace, 2, newMember->name, format1);
				//_mExcelReader->writeExcel(takenSpace + memberSpace, 3, newMember->gender, format1); //int?
				_mExcelReader->writeExcel(takenSpace, 4, newMember->duty, format1);
				_mExcelReader->writeExcel(takenSpace, 5, newMember->dutyClass, format1);
				for (int k = 0; k < patternCount; k++)
				{
					_mExcelReader->writeExcel(takenSpace + k, 6, patternName.at(k), format1);
				}
				_mExcelReader->writeExcel(takenSpace+ patternCount, 6, u8"合计", format1);

				_mExcelReader->mergeCells(takenSpace, 2, takenSpace + patternCount, 2, format1);
				_mExcelReader->mergeCells(takenSpace, 3, takenSpace + patternCount, 3, format1);
				_mExcelReader->mergeCells(takenSpace, 4, takenSpace + patternCount, 4, format1);
				_mExcelReader->mergeCells(takenSpace, 5, takenSpace + patternCount, 5, format1);

				//_info->EvaluationMemberInfo->at(0)->EvaluationMembers->at(i)->EvaluationMembers->at(j)->name;
				takenSpace += patternCount + 1;
			}
			_mExcelReader->mergeCells(takenSpace - memberCount *(patternCount + 1), 1, takenSpace - 1, 1, format1);
		}

		_mExcelReader->chooseSheet(2);
		//QString sheetTitle = 
		_mExcelReader->writeExcel(1, 1, u8"得票情况汇总", format1);
		_mExcelReader->writeExcel(2, 1, u8"姓名", format1);
		_mExcelReader->writeExcel(2, 2, u8"性别", format1);
		_mExcelReader->writeExcel(2, 3, u8"职务", format1);
		_mExcelReader->writeExcel(2, 4, u8"表类型", format1);
		_mExcelReader->writeExcel(2, 5, u8"收回数", format1);
		//index
		indexCountS = 6;
		for (unsigned i = 0; i < indexCountF; i++)
		{
			_mExcelReader->writeExcel(2, indexCountS, indexNameF.at(i), format1);
			//二级指标
			unsigned j;
			for (j = 0; j < indexNode->at(i)->SecondLevelIndex.count(); j++)
			{
				_mExcelReader->writeExcel(4, indexCountS, indexNode->at(i)->SecondLevelIndex.at(j).name, format1);
				indexCountS++;
			}
			if (j == 0)
			{
				_mExcelReader->mergeCells(2, indexCountS, 4, indexCountS, format1);
				indexCountS++;
			}
			else
			{
				_mExcelReader->mergeCells(2, indexCountS - j, 3, indexCountS - 1, format1);
			}
		}

		_mExcelReader->mergeCells(1, 1, 1, indexCountS - 1, format1);
		_mExcelReader->mergeCells(2, 1, 4, 1, format1);
		_mExcelReader->mergeCells(2, 2, 4, 2, format1);
		_mExcelReader->mergeCells(2, 3, 4, 3, format1);
		_mExcelReader->mergeCells(2, 4, 4, 4, format1);
		_mExcelReader->mergeCells(2, 5, 4, 5, format1);
		//index

		//初始化结束

		//填充数据
		QList<int> receiveCount;//收回数 计数patternCount

		for (int resultCount = 0; resultCount < results->count(); resultCount++)
		{
			currentResult = results->at(resultCount);

			//if(currentResult->PatternGUID)
			
			qDebug() << currentResult->PatternGUID;

		}

		/*
		//int excelRow = result->GetFormResultCount();
		//默认一个识别结果里面，都是根据一个模板进行识别的，因此结果的分组数是一样的，取第一个FormResult的分组数
		//int excelColumn = result->FormReuslts->at(0)->MarkGroupResults->count();

		//int patternCount = 0;

		//QString patternName = _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount)->FileName;
		_mExcelReader->chooseSheet(1);
		unsigned k;
		for (unsigned j = 0; j < excelColumn; j++)
		{
			_mExcelReader->writeExcel(1, j + 1, result->FormReuslts->at(0)->MarkGroupResults->at(j)->GroupName, format1);
			k = j;
		}
		_mExcelReader->writeExcel(1, k + 2, result->FormReuslts->at(0)->IdentifierResult->Name, format1);
		if (formPattern->ResultType == NumberValue)
		{
			for (unsigned i = 0; i < excelRow; i++)
			{
				Cm3::FormResult::MFormResult* tempFormResult = result->FormReuslts->at(i);
				unsigned k;
				for (unsigned j = 0; j < excelColumn; j++)
				{
					_mExcelReader->writeExcel(i + 2, j + 1, tempFormResult->MarkGroupResults->at(j)->NumberResult, format1);
					k = j;
				}
				_mExcelReader->writeExcel(i + 2, k + 2, tempFormResult->IdentifierResult->Result, format1);
			}
		}
		else
		{
			for (unsigned i = 0; i < excelRow; i++)
			{
				qDebug() << "excel add" << i << " result";
				qDebug() << "MarkGroupResults count" << result->FormReuslts->at(i)->MarkGroupResults->count() << " result";
				//qDebug() << "excel add" << i << " result";
				Cm3::FormResult::MFormResult* tempFormResult = result->GetFormResult(i);
				unsigned k;
				for (unsigned j = 0; j < excelColumn; j++)
				{
					qDebug() << "MarkGroupResults at" << j;
					_mExcelReader->writeExcel(i + 2, j + 1, tempFormResult->GetGroupResult(j)->TextResult, format1);
					k = j;
					qDebug() << "k  at" << k;
				}
				_mExcelReader->writeExcel(i + 2, k + 2, tempFormResult->IdentifierResult->Result, format1);
			}
		}

		//	填充统计数据

		if (formPattern->ResultType == SingleSelect)
		{
			struct AnalyzeResult //建立一个结构体记录每一个统计对象的计数
			{
				QString patternName;
				int count;
			};
			QList<AnalyzeResult> patternResult;
			QList<QList<AnalyzeResult>> patternResults;
			_mExcelReader->chooseSheet(2);
			patternResults.clear();
			//for (int k = 0; k < excelRow; k++)
			{
				for (int i = 0; i < formPattern->MarkGroupPattern->size(); i++)
				{
					patternResult.clear();
					for (int j = 0; j < formPattern->MarkGroupPattern->at(i)->CellList->size(); j++)
					{
						AnalyzeResult eachPatternResult;
						eachPatternResult.patternName = formPattern->MarkGroupPattern->at(i)->CellList->at(j)->CellName;
						eachPatternResult.count = 0;
						patternResult.append(eachPatternResult);
					}
					//往patternResult里面填充未选和多选两种情况
					AnalyzeResult eachPatternResult1;
					AnalyzeResult eachPatternResult2;
					eachPatternResult1.patternName = QString("*");
					eachPatternResult2.patternName = QString("");
					eachPatternResult1.count = 0;
					eachPatternResult2.count = 0;
					patternResult.append(eachPatternResult1);
					patternResult.append(eachPatternResult2);
					//最后在patternResults中添加当前patternResult
					patternResults.append(patternResult);
				}
			}
			//计算统计用参数
			for (int i = 0; i < result->GetFormResultCount(); i++)
			{
				for (int j = 0; j < result->FormReuslts->at(i)->GetGroupResultCount(); j++)
				{
					for (int k = 0; k < patternResults.at(j).size(); k++)
					{
						if (result->FormReuslts->at(i)->GetGroupResult(j)->TextResult == patternResults.at(j).at(k).patternName)
						{
							AnalyzeResult tempAnalyzeResult;
							tempAnalyzeResult = patternResults.at(j).at(k);
							tempAnalyzeResult.count += 1;
							patternResults[j].replace(k, tempAnalyzeResult);
							//patternResult[k].count++;
						}
					}
				}
			}
			//初始化行与列的索引
			int rowIndex = 1;
			int columnIndex = 1;
			//表头设计，根据得到的patternNames确定每一个group的大小，合并相应长度的单元格
			for (int i = 0; i < patternResults.size(); i++)
			{
				int lenGroup = patternResults.at(i).size();
				_mExcelReader->mergeCells(rowIndex, columnIndex, rowIndex, columnIndex + lenGroup - 1, format1);
				for (int j = 0; j < patternResults.at(i).size(); j++)
				{
					_mExcelReader->writeExcel(rowIndex, columnIndex, formPattern->MarkGroupPattern->at(i)->GroupName, format1);
					QString tempPatternName = patternResults.at(i).at(j).patternName;
					if (tempPatternName == QString("*"))
					{
						_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, (u8"多选"), format1);
					}
					else if (tempPatternName == QString(""))
					{
						_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, (u8"未选"), format1);
					}
					else
					{
						_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, patternResults.at(i).at(j).patternName, format1);
					}
					_mExcelReader->writeExcel(rowIndex + 1 + 1, columnIndex + j, QString::number(patternResults.at(i).at(j).count), format1);
				}
				columnIndex = columnIndex + lenGroup;
			}
		}
	}

	/*
	for (int resultConut = 0; resultConut < _results->count(); resultConut++)
	{
		currentResult = _results->at(resultConut);
		MPattern* pattern = _info->RecognizePatternInfo->RecognizeFormPatterns->at(resultConut);
		for (int patternCount = 0; patternCount < pattern->GetFormPatternCount(); patternCount++)
		{
			MFormPattern* formPattern = pattern->GetFormPattern(patternCount);
			MResult* result = new MResult();
			for (int i = 0; i < currentResult->GetFormResultCount(); i++)
			{
				if (currentResult->GetFormResult(i)->FormIndex == patternCount)
				{
					result->AddFormResult(currentResult->GetFormResult(i));
				}
			}
			if (result->GetFormResultCount() > 0)
			{

				//创建一个改名字的excel表
				//QFile file(excelTempName);
				//file.open(QIODevice::ReadWrite);
				//file.close();
				if (_mExcelReader->newExcel(excelTempName))
				{
					_mExcelReader->setSheetName(evaluationName);

					//填充明细数据
					//单位页
					_mExcelReader->chooseSheet(0);
					_mExcelReader->writeExcel(1, 1, u8"单位指标汇总表", format1);

					_mExcelReader->writeExcel(2, 1, u8"单位", format1);
					_mExcelReader->writeExcel(2, 2, u8"表类型", format1);
					_mExcelReader->writeExcel(2, 3, u8"收回数", format1);

					_mExcelReader->mergeCells(1, 1, 1, 3, format1);
					int memberCount = _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount)->EvaluationUnits.count();
					for (unsigned j = 0; j < memberCount; j++)
					{
						_mExcelReader->writeExcel(2, j + 4, _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount)->EvaluationUnits.at(j), format1);
					}
					int
					//int excelRow = result->GetFormResultCount();
					//默认一个识别结果里面，都是根据一个模板进行识别的，因此结果的分组数是一样的，取第一个FormResult的分组数
					//int excelColumn = result->FormReuslts->at(0)->MarkGroupResults->count();

					//int patternCount = 0;

					QString patternName = _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount)->FileName;
					_mExcelReader->chooseSheet(1);
					unsigned k;
					for (unsigned j = 0; j < excelColumn; j++)
					{
						_mExcelReader->writeExcel(1, j + 1, result->FormReuslts->at(0)->MarkGroupResults->at(j)->GroupName, format1);
						k = j;
					}
					_mExcelReader->writeExcel(1, k + 2, result->FormReuslts->at(0)->IdentifierResult->Name, format1);
					if (formPattern->ResultType == NumberValue)
					{
						for (unsigned i = 0; i < excelRow; i++)
						{
							Cm3::FormResult::MFormResult* tempFormResult = result->FormReuslts->at(i);
							unsigned k;
							for (unsigned j = 0; j < excelColumn; j++)
							{
								_mExcelReader->writeExcel(i + 2, j + 1, tempFormResult->MarkGroupResults->at(j)->NumberResult, format1);
								k = j;
							}
							_mExcelReader->writeExcel(i + 2, k + 2, tempFormResult->IdentifierResult->Result, format1);
						}
					}
					else
					{
						for (unsigned i = 0; i < excelRow; i++)
						{
							qDebug() << "excel add" << i << " result";
							qDebug() << "MarkGroupResults count" << result->FormReuslts->at(i)->MarkGroupResults->count() << " result";
							//qDebug() << "excel add" << i << " result";
							Cm3::FormResult::MFormResult* tempFormResult = result->GetFormResult(i);
							unsigned k;
							for (unsigned j = 0; j < excelColumn; j++)
							{
								qDebug() << "MarkGroupResults at" << j;
								_mExcelReader->writeExcel(i + 2, j + 1, tempFormResult->GetGroupResult(j)->TextResult, format1);
								k = j;
								qDebug() << "k  at" << k;
							}
							_mExcelReader->writeExcel(i + 2, k + 2, tempFormResult->IdentifierResult->Result, format1);
						}
					}

					//	填充统计数据

					if (formPattern->ResultType == SingleSelect)
					{
						struct AnalyzeResult //建立一个结构体记录每一个统计对象的计数
						{
							QString patternName;
							int count;
						};
						QList<AnalyzeResult> patternResult;
						QList<QList<AnalyzeResult>> patternResults;
						_mExcelReader->chooseSheet(2);
						patternResults.clear();
						//for (int k = 0; k < excelRow; k++)
						{
							for (int i = 0; i < formPattern->MarkGroupPattern->size(); i++)
							{
								patternResult.clear();
								for (int j = 0; j < formPattern->MarkGroupPattern->at(i)->CellList->size(); j++)
								{
									AnalyzeResult eachPatternResult;
									eachPatternResult.patternName = formPattern->MarkGroupPattern->at(i)->CellList->at(j)->CellName;
									eachPatternResult.count = 0;
									patternResult.append(eachPatternResult);
								}
								//往patternResult里面填充未选和多选两种情况
								AnalyzeResult eachPatternResult1;
								AnalyzeResult eachPatternResult2;
								eachPatternResult1.patternName = QString("*");
								eachPatternResult2.patternName = QString("");
								eachPatternResult1.count = 0;
								eachPatternResult2.count = 0;
								patternResult.append(eachPatternResult1);
								patternResult.append(eachPatternResult2);
								//最后在patternResults中添加当前patternResult
								patternResults.append(patternResult);
							}
						}
						//计算统计用参数
						for (int i = 0; i < result->GetFormResultCount(); i++)
						{
							for (int j = 0; j < result->FormReuslts->at(i)->GetGroupResultCount(); j++)
							{
								for (int k = 0; k < patternResults.at(j).size(); k++)
								{
									if (result->FormReuslts->at(i)->GetGroupResult(j)->TextResult == patternResults.at(j).at(k).patternName)
									{
										AnalyzeResult tempAnalyzeResult;
										tempAnalyzeResult = patternResults.at(j).at(k);
										tempAnalyzeResult.count += 1;
										patternResults[j].replace(k, tempAnalyzeResult);
										//patternResult[k].count++;
									}
								}
							}
						}
						//初始化行与列的索引
						int rowIndex = 1;
						int columnIndex = 1;
						//表头设计，根据得到的patternNames确定每一个group的大小，合并相应长度的单元格
						for (int i = 0; i < patternResults.size(); i++)
						{
							int lenGroup = patternResults.at(i).size();
							_mExcelReader->mergeCells(rowIndex, columnIndex, rowIndex, columnIndex + lenGroup - 1, format1);
							for (int j = 0; j < patternResults.at(i).size(); j++)
							{
								_mExcelReader->writeExcel(rowIndex, columnIndex, formPattern->MarkGroupPattern->at(i)->GroupName, format1);
								QString tempPatternName = patternResults.at(i).at(j).patternName;
								if (tempPatternName == QString("*"))
								{
									_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, (u8"多选"), format1);
								}
								else if (tempPatternName == QString(""))
								{
									_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, (u8"未选"), format1);
								}
								else
								{
									_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, patternResults.at(i).at(j).patternName, format1);
								}
								_mExcelReader->writeExcel(rowIndex + 1 + 1, columnIndex + j, QString::number(patternResults.at(i).at(j).count), format1);
							}
							columnIndex = columnIndex + lenGroup;
						}
					}
				}
			}
			//在进入下一个循环之前，delete掉新建的指针
			delete result;
			result = nullptr;


		}
	}
	*/
	}

}
