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
/*
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
		for (int ii = 0; ii < lenResults; ii++)
		{
			MResult* rst = _results->at(ii);
			input >> rst->PatternGUID >> rst->_isInitialized >> rst->_validFormCount >> rst->FormResultsLen;
			int len = rst->FormResultsLen;
			for (int i = 0; i < len; i++)
			{
				getCount(i, len);
				MFormResult* mfrst = new MFormResult();
				_results->at(ii)->AddFormResult(mfrst);
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

*/

void LoadDataFile::generateExcelResult(QList<MResult*>* results, QString excelName)
{
	//规定excel的单元格式样式
	QXlsx::Format format1;
	format1.setFontColor(QColor(Qt::black));
	format1.setPatternBackgroundColor(QColor(Qt::white));
	format1.setFontSize(11);
	format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
	format1.setVerticalAlignment(QXlsx::Format::AlignVCenter);
	format1.setBorderStyle(QXlsx::Format::BorderMedium);/*边框样式*/

	MResult* currentResult;
	for (int patternCount = 0; _info->RecognizePatternInfo->RecognizeFormPatterns->count(); patternCount++)
	{
		MRecognizeFormPattern* _currentPattern = _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount);
		QString tableIndex = _currentPattern->GetFormPattern(patternCount)->IdentifierCodePattern->CodeValue;

		//currentResult = results->at(patternCount);
		ES2EvaluationMembers* currentMemberInfo = _info->EvaluationMemberInfo->at(patternCount);
		//直接进行保存
		QString patternName = _currentPattern->FileName;
		QString excelFileName = patternName + ".xlsx";
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
					excelFileName = patternName + "(" + QString::number(repeatCount) + ").xlsx";
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
			_mExcelReader->setSheetName(patternName);

			//初始化表头
			//单位页
			_mExcelReader->chooseSheet(0);
			/*
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
			int unitCount = _currentPattern->EvaluationUnits.count();
			QList<QString> unitName;
			unitName.clear();
			for (unsigned i = 0; i < unitCount; i++)
			{
				unitName.append(_currentPattern->EvaluationUnits.at(i));
				_mExcelReader->writeExcel(i * (patternCount + 1) + 5, 1, unitName.at(i), format1);
				_mExcelReader->mergeCells(i * (patternCount + 1) + 5, 1, (i + 1) * (patternCount + 1) + 4, 1, format1);

			}
			//遍历指标
			int indexCountF = _currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->count();
			int indexCountGroup = 0;
			int indexCountT = 4;
			QList<QString> indexNameF;
			QList<MemberDetailIndex*>* indexNode = new QList<MemberDetailIndex*>;
			indexNameF.clear();
			indexNode->clear();
			for (unsigned i = 0; i < indexCountF; i++)
			{
				indexNode->append(_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at(i));
				//_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at()->FirstLevelIndex.name
				indexNameF.append(indexNode->at(i)->FirstLevelIndex.name);
				_mExcelReader->writeExcel(2, indexCountT, indexNameF.at(i), format1);
				int initMark = indexCountT;
				//二级指标
				unsigned j;
				for (j = 0; j < indexNode->at(i)->SecondLevelIndex.count(); j++)
				{
					_mExcelReader->writeExcel(3, indexCountT, indexNode->at(i)->SecondLevelIndex.at(j).name, format1);
					unsigned k;
					MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(0)->MarkGroupPattern->at(indexCountGroup);
					for (k = 0; k < tempGroupPattern->CellList->count(); k++)
					{
						_mExcelReader->writeExcel(4, indexCountT, tempGroupPattern->CellList->at(k)->CellName, format1);
						indexCountT++;
					}
					indexCountGroup++;
					if (k == 0)
					{
						_mExcelReader->mergeCells(3, indexCountT, 4, indexCountT, format1);
						indexCountT++;
					}
					else
					{
						_mExcelReader->mergeCells(3, indexCountT - k, 3, indexCountT - 1, format1);
					}
				}
				if (j == 0)
				{
					unsigned k;
					MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(0)->MarkGroupPattern->at(indexCountGroup);
					for (k = 0; k < tempGroupPattern->CellList->count(); k++)
					{
						_mExcelReader->writeExcel(4, indexCountT, tempGroupPattern->CellList->at(k)->CellName, format1);
						indexCountT++;
					}
					indexCountGroup++;
					if (k == 0)
					{
						_mExcelReader->mergeCells(2, indexCountT, 4, indexCountT, format1);
						indexCountT++;
					}
					else
					{
						_mExcelReader->mergeCells(2, indexCountT - k, 3, indexCountT - 1, format1);
					}
				}
				else
				{
					_mExcelReader->mergeCells(2, initMark, 2, indexCountT - 1, format1);
				}
			}

			//_mExcelReader->mergeCells(1, 1, 1, 3, format1);//无指标情况
			_mExcelReader->mergeCells(1, 1, 1, indexCountT - 1, format1);
			*/
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
			/*
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
			*/
			int indexCountF = _currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->count();
			int indexCountGroup = 0;
			int indexCountT = 8;
			QList<QString> indexNameF;
			QList<MemberDetailIndex*>* indexNode = new QList<MemberDetailIndex*>;
			indexNameF.clear();
			indexNode->clear();
			QList<int> cellListCount;
			for (unsigned i = 0; i < indexCountF; i++)
			{
				indexNode->append(_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at(i));
				//_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at()->FirstLevelIndex.name
				indexNameF.append(indexNode->at(i)->FirstLevelIndex.name);
				_mExcelReader->writeExcel(2, indexCountT, indexNameF.at(i), format1);
				int initMark = indexCountT;
				//二级指标
				unsigned j;
				for (j = 0; j < indexNode->at(i)->SecondLevelIndex.count(); j++)
				{
					_mExcelReader->writeExcel(3, indexCountT, indexNode->at(i)->SecondLevelIndex.at(j).name, format1);
					unsigned k;
					MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(patternCount)->MarkGroupPattern->at(indexNode->at(i)->SecondLevelIndex.at(j).groupIndex);
					cellListCount.append(0);
					cellListCount[indexCountGroup] = tempGroupPattern->CellList->count();
					for (k = 0; k < tempGroupPattern->CellList->count(); k++)
					{
						_mExcelReader->writeExcel(4, indexCountT, tempGroupPattern->CellList->at(k)->CellName, format1);
						indexCountT++;
					}
					indexCountGroup++;
					if (k == 0)
					{
						_mExcelReader->mergeCells(3, indexCountT, 4, indexCountT, format1);
						indexCountT++;
					}
					else
					{
						_mExcelReader->mergeCells(3, indexCountT - k, 3, indexCountT - 1, format1);
					}
				}
				if (j == 0)
				{
					unsigned k;
					MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(patternCount)->MarkGroupPattern->at(indexNode->at(i)->FirstLevelIndex.groupIndex);
					cellListCount.append(0);
					cellListCount[indexCountGroup] = tempGroupPattern->CellList->count();
					for (k = 0; k < tempGroupPattern->CellList->count(); k++)
					{
						_mExcelReader->writeExcel(4, indexCountT, tempGroupPattern->CellList->at(k)->CellName, format1);
						indexCountT++;
					}
					indexCountGroup++;
					if (k == 0)
					{
						_mExcelReader->mergeCells(2, indexCountT, 4, indexCountT, format1);
						indexCountT++;
					}
					else
					{
						_mExcelReader->mergeCells(2, indexCountT - k, 3, indexCountT - 1, format1);
					}
				}
				else
				{
					_mExcelReader->mergeCells(2, initMark, 2, indexCountT - 1, format1);
				}
			}

			_mExcelReader->mergeCells(1, 1, 1, indexCountT - 1, format1);
			_mExcelReader->mergeCells(2, 1, 4, 1, format1);
			_mExcelReader->mergeCells(2, 2, 4, 2, format1);
			_mExcelReader->mergeCells(2, 3, 4, 3, format1);
			_mExcelReader->mergeCells(2, 4, 4, 4, format1);
			_mExcelReader->mergeCells(2, 5, 4, 5, format1);
			_mExcelReader->mergeCells(2, 6, 4, 6, format1);
			_mExcelReader->mergeCells(2, 7, 4, 7, format1);


			int unitCount = _currentPattern->EvaluationUnits.count();
			int subjectCount = 2;
			//_currentPattern->EvaluationSubjectIndex;
		//_info->EvaluationSubjectInfo->EvaluationSubjects->at();
		//_info->RecognizePatternInfo
			//_results;


			QList<QString> unitName;
			unitName.clear();
			for (unsigned i = 0, takenSpace = 5; i < unitCount; i++)
			{
				unitName.append(_currentPattern->EvaluationUnits.at(i));
				_mExcelReader->writeExcel(i * (subjectCount + 1) + 5, 1, unitName.at(i), format1);
				//_mExcelReader->mergeCells(i * (subjectCount + 1) + 5, 1, (i + 1) * (subjectCount + 1) + 4, 1, format1);

				int memberCount = currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->count();
				_mExcelReader->writeExcel(takenSpace, 1, unitName.at(i), format1);
				for (int j = 0; j < memberCount; j++)
				{
					MemberInfo* newMember = currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j);
					_mExcelReader->writeExcel(takenSpace, 2, newMember->name, format1);
					_mExcelReader->writeExcel(takenSpace, 3, (newMember->gender == 0 ? u8"男" : u8"女"), format1); //int 0male/1female
					_mExcelReader->writeExcel(takenSpace, 4, newMember->duty, format1);
					_mExcelReader->writeExcel(takenSpace, 5, newMember->dutyClass, format1);
					for (int k = 0; k < subjectCount; k++)
					{
						_mExcelReader->writeExcel(takenSpace + k, 6, "空数据", format1);//!!!!!!!!!!!!
					}
					_mExcelReader->writeExcel(takenSpace + subjectCount, 6, u8"合计", format1);

					_mExcelReader->mergeCells(takenSpace, 2, takenSpace + subjectCount, 2, format1);
					_mExcelReader->mergeCells(takenSpace, 3, takenSpace + subjectCount, 3, format1);
					_mExcelReader->mergeCells(takenSpace, 4, takenSpace + subjectCount, 4, format1);
					_mExcelReader->mergeCells(takenSpace, 5, takenSpace + subjectCount, 5, format1);

					//currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j)->name;
					takenSpace += subjectCount + 1;
				}
				_mExcelReader->mergeCells(takenSpace - memberCount * (subjectCount + 1), 1, takenSpace - 1, 1, format1);
			}
			/*
						_mExcelReader->chooseSheet(2);
						//QString sheetTitle =
						_mExcelReader->writeExcel(1, 1, u8"得票情况汇总", format1);
						_mExcelReader->writeExcel(2, 1, u8"姓名", format1);
						_mExcelReader->writeExcel(2, 2, u8"性别", format1);
						_mExcelReader->writeExcel(2, 3, u8"职务", format1);
						_mExcelReader->writeExcel(2, 4, u8"表类型", format1);
						_mExcelReader->writeExcel(2, 5, u8"收回数", format1);

						indexCountT = 6;
						indexCountGroup = 0;
						for (unsigned i = 0; i < indexCountF; i++)
						{
							indexNode->append(_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at(i));
							//_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at()->FirstLevelIndex.name
							indexNameF.append(indexNode->at(i)->FirstLevelIndex.name);
							_mExcelReader->writeExcel(2, indexCountT, indexNameF.at(i), format1);
							int initMark = indexCountT;
							//二级指标
							unsigned j;
							for (j = 0; j < indexNode->at(i)->SecondLevelIndex.count(); j++)
							{
								_mExcelReader->writeExcel(3, indexCountT, indexNode->at(i)->SecondLevelIndex.at(j).name, format1);
								unsigned k;
								MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(0)->MarkGroupPattern->at(indexCountGroup);
								for (k = 0; k < tempGroupPattern->CellList->count(); k++)
								{
									_mExcelReader->writeExcel(4, indexCountT, tempGroupPattern->CellList->at(k)->CellName, format1);
									indexCountT++;
								}
								indexCountGroup++;
								if (k == 0)
								{
									_mExcelReader->mergeCells(3, indexCountT, 4, indexCountT, format1);
									indexCountT++;
								}
								else
								{
									_mExcelReader->mergeCells(3, indexCountT - k, 3, indexCountT - 1, format1);
								}
							}
							if (j == 0)
							{
								unsigned k;
								MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(0)->MarkGroupPattern->at(indexCountGroup);
								for (k = 0; k < tempGroupPattern->CellList->count(); k++)
								{
									_mExcelReader->writeExcel(4, indexCountT, tempGroupPattern->CellList->at(k)->CellName, format1);
									indexCountT++;
								}
								indexCountGroup++;
								if (k == 0)
								{
									_mExcelReader->mergeCells(2, indexCountT, 4, indexCountT, format1);
									indexCountT++;
								}
								else
								{
									_mExcelReader->mergeCells(2, indexCountT - k, 3, indexCountT - 1, format1);
								}
							}
							else
							{
								_mExcelReader->mergeCells(2, initMark, 2, indexCountT - 1, format1);
							}
						}

						_mExcelReader->mergeCells(1, 1, 1, indexCountT - 1, format1);
						_mExcelReader->mergeCells(2, 1, 4, 1, format1);
						_mExcelReader->mergeCells(2, 2, 4, 2, format1);
						_mExcelReader->mergeCells(2, 3, 4, 3, format1);
						_mExcelReader->mergeCells(2, 4, 4, 4, format1);
						_mExcelReader->mergeCells(2, 5, 4, 5, format1);
						//index
			*/
			//初始化结束

			//统计数据
			//MRecognizeFormPattern* currentFormPattern;
			QList<QList<QList<QList<QString>>>> scoreCount;//单位-主体-成员-得分情况
			//vector<vector<vector<vector<int>>>> scoreCount;
			QList<int> receiveCount;//收回数 计数subjectIndex

			for (int resultCount = 0; resultCount < results->count(); resultCount++)
			{
				//currentFormPattern = nullptr;
				currentResult = results->at(resultCount);
				
				if (currentResult->FormReuslts->at(0)->IdentifierResult->Result.at(0) == tableIndex) //判断模式一致
				{
					QList<MFormResult*>* currentFormResults = new QList<MFormResult*>;
					for (int unitIndex = 0; unitIndex < _currentPattern->EvaluationUnits.count(); unitIndex++)
					{
						QList<QList<QList<QString>>> scoreCount_unit;
						//scoreCount.append(scoreCount_unit);
						currentFormResults->clear();
						for (int formIndex = 0; formIndex < currentResult->FormReuslts->count(); formIndex++)
						{
							if (currentResult->FormReuslts->at(formIndex)->IsRecognizeSuccess)
							{
								if (currentResult->FormReuslts->at(formIndex)->IdentifierResult->Result.at(6).toLatin1() - 48 == unitIndex)
								{
									currentFormResults->append(currentResult->FormReuslts->at(formIndex));
								}
							}
						}
						for (int subjectIndex = 0; subjectIndex < subjectCount; subjectIndex++)//主体组
						{
							receiveCount.append(0);
							QList<QList<QString>> scoreCount_subject;
							//scoreCount_unit.append(scoreCount_subject);
							for (int formIndex = 0; formIndex < currentFormResults->count(); formIndex++)
							{
								if (currentFormResults->at(formIndex)->IdentifierResult->Result.at(4).toLatin1() - 48 == subjectIndex)
								{
									receiveCount[subjectIndex] ++;
									MFormResult* currentFormResult = currentFormResults->at(formIndex);
									int page = currentFormResult->IdentifierResult->Result.at(2).toLatin1() - 48;
									int shift = 0;
									for (int i = 0; i < page; i++)
									{
										shift += _currentPattern->FormPatternHash.at(i).size();
									}
									int memberCount = _currentPattern->MemberIndexs->count();
									//int memberCount = _currentPattern->FormPatternHash.at(page).size(); //当前页面人数
									QList<QString> scoreCount_member;
									//bool flag = 0, newflag = 1;
									for (int memberIndex = 0; memberIndex <  _currentPattern->FormPatternHash.at(page).size(); memberIndex++)
									{
										scoreCount_member.clear();
										MemberIndex* currentMember = _currentPattern->MemberIndexs->at(memberIndex + shift);
										for (int groupCount = 0; groupCount < indexCountGroup;)
										{
											MemberDetailIndex* currentIndex =  currentMember->MemberDetailIndexs->at(groupCount);
											if (currentIndex->IndexLevel == 1)
											{
												scoreCount_member.append("");
												MGroupResult* currentGroupResult = currentFormResult->MarkGroupResults->at(currentIndex->FirstLevelIndex.groupIndex);
												int score;
												if (scoreCount_subject.size() < memberCount)
												{
													score = 0;
												}
												else
												{
													score = scoreCount_subject[memberIndex + shift][groupCount].toInt();
												}
												score += currentGroupResult->NumberResult.toInt();
												scoreCount_member[groupCount] = QString::number(score);
												groupCount++;
											}
											else if(currentIndex->IndexLevel == 2)
											{
												for (int i = 0; i < currentIndex->SecondLevelIndex.count(); i++)
												{
													scoreCount_member.append("");
													MGroupResult* currentGroupResult = currentFormResult->MarkGroupResults->at(currentIndex->SecondLevelIndex.at(i).groupIndex);
													int score;
													if (scoreCount_subject.size() < memberCount)
													{
														score = 0;
													}
													else
													{
														score = scoreCount_subject[memberIndex + shift][groupCount].toInt();
													}
													score += currentGroupResult->NumberResult.toInt();
													scoreCount_member[groupCount] = QString::number(score);
													groupCount++;
												}
											}
										}
										if (scoreCount_subject.size() >= memberCount)
										{
											scoreCount_subject[memberIndex + shift] = scoreCount_member;
										}
										else
										{
											scoreCount_subject.append(scoreCount_member);
										}
										//currentMember->MemberDetailIndexs->at(0);
									}
									/*
									for (int i = 0; i < currentFormResult->MarkGroupResults->count(); i++)//
									{
										flag = 0;
										MGroupResult* currentGroupResult = currentFormResult->MarkGroupResults->at(i);
										if (memberCount != i / indexCountGroup + shift)
										{
											memberCount = i / indexCountGroup + shift;
											flag = 1;
											if (!newflag)
											{
												scoreCount_subject.append(scoreCount_member);
											}
											else
											{
												newflag = 0;
											}
											scoreCount_member.clear();
										}

										int groupCount = i % indexCountGroup;
										scoreCount_member.append("");

										
										if (scoreCount[unitIndex][subjectIndex][memberCount][groupCount] == nullptr)
										{
											scoreCount[unitIndex][subjectIndex][memberCount][groupCount] = "";
										}
										int score = scoreCount[subjectIndex][unitIndex][memberCount][groupCount].toInt();
										score += currentGroupResult->NumberResult.toInt();
										scoreCount[subjectCount][unitIndex][memberCount][groupCount] = QString::number(score);
										//scoreCount[subjectCount][unitCount][memberCount] += "";

										if (flag) {
											QList<QString> scoreCount_member;
											scoreCount_subject.append(scoreCount_member);
										}
										
										int score = scoreCount_member[groupCount].toInt();
										score += currentGroupResult->NumberResult.toInt();
										scoreCount_member[groupCount] = QString::number(score);
										
										int score = scoreCount_subject[memberCount][groupCount].toInt();
										score += currentGroupResult->NumberResult.toInt();
										scoreCount_subject[memberCount][groupCount] = QString::number(score);
										
									}
									if (true)
									{
										scoreCount_subject.append(scoreCount_member);
										int a = 0;
									}
									*/
								}
							}

							scoreCount_unit.append(scoreCount_subject);

						}
						scoreCount.append(scoreCount_unit);
					}
				}
			}
			scoreCount;
			int pageCount = _currentPattern->GetFormPatternCount();
			receiveCount.append(0);
			for (int i = 0; i < receiveCount.count() - 1; i++)
			{
				receiveCount[i] /= pageCount;
				receiveCount[receiveCount.count() - 1] += receiveCount.at(i);
			}


				//QString resultMark = currentResult->FormReuslts->at(0)->IdentifierResult->Result;
				//int subjectMark = resultMark.at(4).toLatin1() - 48;//主体
				//int unitMark = resultMark.at(6).toLatin1() - 48;//单位


			//输出
			for (int unitIndex = 0; unitIndex < unitCount; unitIndex++)
			{
				int pageCount = _currentPattern->GetFormPatternCount();
				int memberSum = currentMemberInfo->EvaluationMembers->at(unitIndex)->EvaluationMembers->count();
				int excelRowCount = 5;
				int excelColumnIndex = 8;
				int subjectCount = 2;
				
				for (int subjectIndex = 0; subjectIndex < subjectCount; subjectIndex++)
				{
					if (subjectIndex != 0)
					{
						excelRowCount -= memberSum * (subjectCount + 1);
					}
					//_mExcelReader->writeExcel(excelRowCount + subjectCount, 7, QString::number(ceil(receiveCount[subjectCount] * 1.0 / pageCount * 1.0)), format1);
					for (int memberIndex = 0; memberIndex < memberSum; memberIndex++)
					{
						QList<int> memberScore;
						memberScore.append(0);
						_mExcelReader->writeExcel(excelRowCount + subjectIndex, 7, QString::number(receiveCount.at(subjectIndex)), format1);
						for (int i = 0; i < indexCountGroup; i++)
						{
							//MFormPattern* currentformpattern = _currentPattern->GetFormPattern(0);
							//_currentPattern->MemberIndexs->at(0)->MemberDetailIndexs->at(i)->FirstLevelIndex.groupIndex;
							int empty = cellListCount.at(i) - scoreCount[unitIndex][subjectIndex][memberIndex][i].count();
							for (int j = 0; j < empty; j++)
							{
								_mExcelReader->writeExcel(excelRowCount + subjectIndex, excelColumnIndex, "0", format1);
								memberScore[excelColumnIndex - 8] += 0;
								memberScore.append(0);
								excelColumnIndex++;
							}
							for (int n = 0; n < cellListCount.at(i) - empty; n++)
							{
								QString result = scoreCount[unitIndex][subjectIndex][memberIndex][i][n];
								_mExcelReader->writeExcel(excelRowCount + subjectIndex, excelColumnIndex, result, format1);
								memberScore[excelColumnIndex - 8] += result.toInt(); //?
								memberScore.append(0);
								excelColumnIndex++;
							}
						}
						excelColumnIndex = 8;
						if(subjectIndex == subjectCount - 1)
						{
							_mExcelReader->writeExcel(excelRowCount - 1, 7, QString::number(receiveCount.at(subjectCount)), format1);
							for (int i = 0; i < memberScore.size() - 1; i++) //总计
							{
								_mExcelReader->writeExcel(excelRowCount + subjectCount, excelColumnIndex, QString::number(memberScore.at(i)), format1);
								excelColumnIndex++;
							}
						}
						excelColumnIndex = 8;
						excelRowCount += subjectCount + 1;
					}
				}
			
				//resultMark.at(0);//模式
				/*
				if (resultMark.at(0) != QString::number(patternCount))
				{
					continue;
				}
				else
				{
					int subjectMark = resultMark.at(4).toLatin1() - 48;//主体
					int unitMark = resultMark.at(6).toLatin1() - 48;//单位
					int memberCount = currentMemberInfo->EvaluationMembers->at(unitMark)->EvaluationMembers->count();
					for (int i = 0; i < memberCount; i++)
					{

					}
					//currentResult->FormReuslts->

				}


				QString patternID = currentResult->PatternGUID;
				for (int i = 0; i < patternCount; i++)
				{
					currentFormPattern = _info->RecognizePatternInfo->RecognizeFormPatterns->at(i);
					if (patternID == currentFormPattern->GetPatternGUID())
					{
						break;
					}
				}

				if (currentFormPattern == nullptr)
				{
					return;
				}

				for (int formResultCount = 0; formResultCount < currentResult->FormReuslts->count(); formResultCount++)
				{
					//分码1-0-0-0
					MFormResult* tempResult = currentResult->GetFormResult(formResultCount);
					MFormPattern* tempPattern = currentFormPattern->GetFormPattern(formResultCount);
					vector<int> memberIndexsList = currentFormPattern->FormPatternHash[0];
					//
					//
					indexCountGroup;
					_mExcelReader->chooseSheet(1);
					for (int i = 0; i < currentFormPattern->MemberIndexs->count(); i++)
					{
						MemberIndex* currentMember = new MemberIndex();
						currentMember = currentFormPattern->MemberIndexs->at(i);
						for (int j = 0; j < indexCountGroup; j++)
						{
							indexCountT = 8;
							tempPattern->MarkGroupPattern->at(j + i * indexCountGroup);
							//_mExcelReader->writeExcel(, indexCountT, );
							//float modulus = currentDetailIndex->FirstLevelIndex.weightNumerator / currentDetailIndex->FirstLevelIndex.weightDenominator;

							int score = currentResult->FormReuslts->at(patternCount)->MarkGroupResults->at(i)->NumberResult.toInt();//
						}


						//取得分数
					}
				}
				*/
				//取得对应下标
				//currentResult->FormReuslts->at(0)->MarkGroupResults->at(i)->
				//取得主体权重

				//
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
			*/
			//	填充统计数据
			/*
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

				//在进入下一个循环之前，delete掉新建的指针
				delete result;
				result = nullptr;


			}
		}
		*/
		}

	}

}
