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
	_currentPattern = new MRecognizeFormPattern();
	_currentMemberInfo = new ES2EvaluationMembers();
}

LoadDataFile::~LoadDataFile()
{
	if (_mExcelReader != nullptr)
	{
		delete _mExcelReader;
		_mExcelReader = nullptr;
	}
	if (_info != nullptr)
	{
		delete _info;
		_info = nullptr;
	}
	if (_currentPattern != nullptr)
	{
		delete _currentPattern;
		_currentPattern = nullptr;
	}	if (_currentMemberInfo != nullptr)
	{
		delete _currentMemberInfo;
		_currentMemberInfo = nullptr;
	}

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

void LoadDataFile::setExcelData(MResult* result, QString excelName, int outputType)
{
	_result = result;
	_excelName = excelName;
	_outputType = outputType;
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
		if (_outputType == 1)
		{
			generateExcelResult(_results, _excelName);
		}
		else if(_outputType == 0)
		{
			generateTestResult(_results, _excelName);
		}
		else
		{
			emit outputError();
			return;
		}

		emit finishExcel();
	}
    else
    {
		emit outputError();
    }
}

void LoadDataFile::setAnswer(QList<QString> answerList)
{
	_answerList = answerList;
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

void LoadDataFile::formatSet(QXlsx::Format &format)
{
	//规定excel的单元格式样式
	//QXlsx::Format format1;
	format.setFontColor(QColor(Qt::black));
	format.setPatternBackgroundColor(QColor(Qt::white));
	format.setFontSize(11);
	format.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
	format.setVerticalAlignment(QXlsx::Format::AlignVCenter);
	format.setBorderStyle(QXlsx::Format::BorderMedium);/*边框样式*/
}

void LoadDataFile::initSheet(int sheetIndex, int patternCount)
{
	//sheetIndex = memberType;
	_mExcelReader->chooseSheet(sheetIndex);
	_mExcelReader->writeExcel(1, 1, u8"成员指标汇总表", format1);
	_mExcelReader->writeExcel(2, 1, u8"单位", format1);
	_mExcelReader->writeExcel(2, 2, u8"姓名", format1);
	_mExcelReader->writeExcel(2, 3, u8"性别", format1);
	_mExcelReader->writeExcel(2, 4, u8"职务", format1);
	_mExcelReader->writeExcel(2, 5, u8"职务类型", format1);
	_mExcelReader->writeExcel(2, 6, u8"表类型", format1);
	_mExcelReader->writeExcel(2, 7, u8"权重", format1);
	_mExcelReader->writeExcel(2, 8, u8"收回数", format1);

	int indexCountF = _currentPattern->MemberIndexs->at(0)->MemberDetailIndexs->count();
	indexCountGroup = 0;
	int indexCountT = 9;
	QList<QString> indexNameF;
	QList<MemberDetailIndex*>* indexNode = new QList<MemberDetailIndex*>;
	indexNameF.clear();
	indexNode->clear();
	//cellListCount;	//clear 考虑有效性
	
	for (int memberIndexCount = 0; memberIndexCount < _currentPattern->MemberIndexs->count(); memberIndexCount++)
	{
		if (_currentPattern->MemberIndexs->at(memberIndexCount)->MemberType != sheetIndex)
		{
			continue;
		}
		for (unsigned i = 0; i < indexCountF; i++)
		{
			indexNode->append(_currentPattern->MemberIndexs->at(memberIndexCount)->MemberDetailIndexs->at(i));
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
				MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(0)->MarkGroupPattern->at(indexNode->at(i)->SecondLevelIndex.at(j).groupIndex);
				cellListCount[sheetIndex].append(0);
				cellListCount[sheetIndex][indexCountGroup] = tempGroupPattern->CellList->count();
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
				MGroupPattern* tempGroupPattern = _currentPattern->GetFormPattern(0)->MarkGroupPattern->at(indexNode->at(i)->FirstLevelIndex.groupIndex);
				cellListCount[sheetIndex].append(0);
				cellListCount[sheetIndex][indexCountGroup] = tempGroupPattern->CellList->count();
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

	}

	_mExcelReader->mergeCells(1, 1, 1, indexCountT - 1, format1);
	_mExcelReader->mergeCells(2, 1, 4, 1, format1);
	_mExcelReader->mergeCells(2, 2, 4, 2, format1);
	_mExcelReader->mergeCells(2, 3, 4, 3, format1);
	_mExcelReader->mergeCells(2, 4, 4, 4, format1);
	_mExcelReader->mergeCells(2, 5, 4, 5, format1);
	_mExcelReader->mergeCells(2, 6, 4, 6, format1);
	_mExcelReader->mergeCells(2, 7, 4, 7, format1);
	_mExcelReader->mergeCells(2, 8, 4, 8, format1);


	unitCount = _currentPattern->EvaluationUnits.count();
	subjectCount = _currentSubject->EvaluationSubjects->count();
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

		int memberCount = _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->count();
		_mExcelReader->writeExcel(takenSpace, 1, unitName.at(i), format1);
		for (int j = 0; j < memberCount; j++)
		{
			MemberInfo* newMember = _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j);
			if (newMember->MemberType != sheetIndex)
			{
				continue;
			}
			_mExcelReader->writeExcel(takenSpace, 2, newMember->name, format1);
			_mExcelReader->writeExcel(takenSpace, 3, (newMember->gender == 0 ? u8"男" : u8"女"), format1); //int 0male/1female
			_mExcelReader->writeExcel(takenSpace, 4, newMember->duty, format1);
			_mExcelReader->writeExcel(takenSpace, 5, newMember->dutyClass, format1);
			for (int k = 0; k < subjectCount; k++)
			{
				_mExcelReader->writeExcel(takenSpace + k, 6, _currentSubject->EvaluationSubjects->at(k)->name, format1);//!!!!!!!!!!!!
			}
			_mExcelReader->writeExcel(takenSpace + subjectCount, 6, u8"合计", format1);

			_mExcelReader->mergeCells(takenSpace, 2, takenSpace + subjectCount, 2, format1);
			_mExcelReader->mergeCells(takenSpace, 3, takenSpace + subjectCount, 3, format1);
			_mExcelReader->mergeCells(takenSpace, 4, takenSpace + subjectCount, 4, format1);
			_mExcelReader->mergeCells(takenSpace, 5, takenSpace + subjectCount, 5, format1);

			//_currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j)->name;
			takenSpace += subjectCount + 1;
		}
		_mExcelReader->mergeCells(takenSpace - memberCount * (subjectCount + 1), 1, takenSpace - 1, 1, format1);
	}
}

void LoadDataFile::initSheet_test(int sheetIndex, int patternCount)
{
	_mExcelReader->chooseSheet(sheetIndex);
	_mExcelReader->writeExcel(1, 1, u8"成员指标汇总表", format1);
	_mExcelReader->writeExcel(2, 1, u8"单位", format1);
	_mExcelReader->writeExcel(2, 2, u8"姓名", format1);
	_mExcelReader->writeExcel(2, 3, u8"性别", format1);
	_mExcelReader->writeExcel(2, 4, u8"职务", format1);
	_mExcelReader->writeExcel(2, 5, u8"职务类型", format1);
	_mExcelReader->writeExcel(2, 6, u8"总得分", format1);
	_mExcelReader->writeExcel(2, 7, u8"明细", format1);

	indexCountGroup = 0;
	int indexCountT = 8;
	QList<QString> indexNameF;
	QList<MemberDetailIndex*>* indexNode = new QList<MemberDetailIndex*>;
	indexNameF.clear();
	indexNode->clear();
	for (int pageIndex = 0; pageIndex < _currentPattern->MemberIndexs->count(); pageIndex++)
	{
		for (unsigned i = 0; i < _currentPattern->MemberIndexs->at(pageIndex)->MemberDetailIndexs->count(); i++)
		{
			indexNode->append(_currentPattern->MemberIndexs->at(pageIndex)->MemberDetailIndexs->at(i));
			//_currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->at()->FirstLevelIndex.name
			indexNameF.append(indexNode->last()->SecondIndexName);
			_mExcelReader->writeExcel(2, indexCountT, indexNameF.last(), format1);
			int initMark = indexCountT;
			//二级指标
			unsigned j;
			for (j = 0; j < indexNode->last()->SecondLevelIndex.count(); j++)
			{
				_mExcelReader->writeExcel(4, indexCountT, indexNode->last()->SecondLevelIndex.at(j).name, format1);
				indexCountT++;
			}
			if (j == 0)
			{
				_mExcelReader->mergeCells(2, indexCountT, 4, indexCountT, format1);
				indexCountT++;
			}
			else
			{
				_mExcelReader->mergeCells(2, indexCountT - j, 3, indexCountT - 1, format1);
			}
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


	unitCount = _currentPattern->EvaluationUnits.count();
	//_currentPattern->EvaluationSubjectIndex;
	//_info->EvaluationSubjectInfo->EvaluationSubjects->at();
	//_info->RecognizePatternInfo
	//_results;


	QList<QString> unitName;
	unitName.clear();
	for (unsigned i = 0, takenSpace = 5; i < unitCount; i++)
	{
		unitName.append(_currentPattern->EvaluationUnits.at(i));
		_mExcelReader->writeExcel(i * 4 + 5, 1, unitName.at(i), format1);
		//_mExcelReader->mergeCells(i * 4 + 5, 1, (i + 1) * 4 + 4, 1, format1);

		int memberCount = _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->count();
		_mExcelReader->writeExcel(takenSpace, 1, unitName.at(i), format1);
		for (int j = 0; j < memberCount; j++)
		{
			MemberInfo* newMember = _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j);
			_mExcelReader->writeExcel(takenSpace, 2, newMember->name, format1);
			_mExcelReader->writeExcel(takenSpace, 3, (newMember->gender == 0 ? u8"男" : u8"女"), format1); //int 0male/1female
			_mExcelReader->writeExcel(takenSpace, 4, newMember->duty, format1);
			_mExcelReader->writeExcel(takenSpace, 5, newMember->dutyClass, format1);

			_mExcelReader->writeExcel(takenSpace + 0, 7, u8"填涂结果", format1);
			_mExcelReader->writeExcel(takenSpace + 1, 7, u8"正确结果", format1);
			_mExcelReader->writeExcel(takenSpace + 2, 7, u8"判题结果", format1);
			_mExcelReader->writeExcel(takenSpace + 3, 7, u8"得分", format1);

			_mExcelReader->mergeCells(takenSpace, 2, takenSpace + 3, 2, format1);
			_mExcelReader->mergeCells(takenSpace, 3, takenSpace + 3, 3, format1);
			_mExcelReader->mergeCells(takenSpace, 4, takenSpace + 3, 4, format1);
			_mExcelReader->mergeCells(takenSpace, 5, takenSpace + 3, 5, format1);
			_mExcelReader->mergeCells(takenSpace, 6, takenSpace + 3, 6, format1);


			//_currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j)->name;
			takenSpace += 3 + 1;
		}
		_mExcelReader->mergeCells(takenSpace - memberCount * 4, 1, takenSpace - 1, 1, format1);
	}
}

void LoadDataFile::generateTestResult(QList<MResult*>* results, QString excelName)
{
	formatSet(format1);
	formatSet(format2);
	format2.setFontColor(QColor(Qt::red));

	for (int patternCount = 0; patternCount < _info->RecognizePatternInfo->RecognizeFormPatterns->count(); patternCount++)
	{
		_currentPattern = _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount);
		QString tableIndex = _currentPattern->GetFormPattern(0)->IdentifierCodePattern->CodeValue;
		//_currentSubject = _info->EvaluationSubjectInfo->EvaluationSubjects->at(_currentPattern->EvaluationSubjectIndex);
		_currentMemberInfo = _info->EvaluationMemberInfo->at(patternCount);
		//直接进行保存
		QString patternName = _currentPattern->FileName;
		QString excelFileName = patternName + ".xlsx";
		//如果文件名重复
		QDir* dir = new QDir(excelName);
		QList<QFileInfo>* fileInfo = new QList<QFileInfo>(dir->entryInfoList());
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
			//初始化
			QList<QString> sheetName;
			sheetName.append(patternName);
			_mExcelReader->setSheetName(sheetName);
			initSheet_test(0, patternCount);


			//数据统计
			QList<QList<QString>> resultCollect;//人员memberIndex / 题号groupIndex
			QList<QString> tempString;
			for (int i = 0; i < _currentMemberInfo->EvaluationMembers->count(); i++)
			{
				for (int j = 0; j < _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->count(); j++)
				{
					resultCollect.append(tempString);
				}
			}
			QList<int> scoreWeight; //分值
			int scoreSum = 0; //add?
			for (int resultIndex = 0; resultIndex < results->count(); resultIndex++)
			{
				_currentResult = results->at(resultIndex);
				if (_currentResult->FormReuslts->at(0)->IdentifierResult->Result.at(0) == tableIndex)
				{
					for (int formCount = 0; formCount < _currentResult->FormReuslts->count(); formCount++)
					{
						MFormResult* currentFormResult = _currentResult->FormReuslts->at(formCount);
						int memberIndex = currentFormResult->IdentifierResult->Result.at(6).toLatin1() - 48;
						int pageIndex = currentFormResult->IdentifierResult->Result.at(2).toLatin1() - 48;
						//QList<QString> resultCollect_member;
						//int groupIndex = 0; 
						//resultCollect_member.append("");

						MemberIndex* currentMemberIndex = _currentPattern->MemberIndexs->at(pageIndex);
						for (int i = 0; i < currentMemberIndex->MemberDetailIndexs->count(); i++)
						{
							MemberDetailIndex* currentMemberDetail = currentMemberIndex->MemberDetailIndexs->at(i);
							for (int j = 0; j < currentMemberDetail->SecondLevelIndex.count(); j++)
							{
								resultCollect[memberIndex].append(currentFormResult->MarkGroupResults->at(currentMemberDetail->SecondLevelIndex.at(j).groupIndex)->TextResult);
								if (memberIndex == 0)
								{
									scoreWeight.append(currentMemberDetail->SecondLevelIndex.at(j).weightNumerator);
									scoreSum += scoreWeight.last();
								}
							}
						}
						//resultCollect[memberIndex].append(resultCollect_member);
					}
				}
			}
			resultCollect;

			//结果输出
			int pageCount = _currentPattern->GetFormPatternCount();
			int rowCount = 5;
			int colunmCount = 8;
			for (int memberIndex = 0; memberIndex < resultCollect.size(); memberIndex++)
			{
				colunmCount = 8;
				int scoreGet = 0;
				for (int groupIndex = 0; groupIndex < resultCollect.at(memberIndex).size(); groupIndex++)
				{
					_mExcelReader->writeExcel(rowCount, colunmCount, resultCollect.at(memberIndex).at(groupIndex), format1);
					_mExcelReader->writeExcel(rowCount + 1, colunmCount, _answerList.at(groupIndex), format1);
					QString score = QString::number(scoreWeight.at(groupIndex));
					if (resultCollect.at(memberIndex).at(groupIndex) == _answerList.at(groupIndex))
					{
						_mExcelReader->writeExcel(rowCount + 2, colunmCount, u8"对", format1);
						_mExcelReader->writeExcel(rowCount + 3, colunmCount, score + "/" + score, format1);
						scoreGet += score.toInt();
					}
					else
					{
						_mExcelReader->writeExcel(rowCount + 2, colunmCount, u8"错", format2);
						_mExcelReader->writeExcel(rowCount + 3, colunmCount, "0/" + score, format1);
					}
					colunmCount ++;
				}
				_mExcelReader->writeExcel(rowCount, 6, QString::number(scoreGet) + "/" + QString::number(scoreSum), format1);
				rowCount += 4;
			}
		}
	}
}

void LoadDataFile::generateExcelResult(QList<MResult*>* results, QString excelName)
{
	formatSet(format1);

	for (int patternCount = 0; patternCount < _info->RecognizePatternInfo->RecognizeFormPatterns->count(); patternCount++)
	{
		_currentPattern = _info->RecognizePatternInfo->RecognizeFormPatterns->at(patternCount);
		QString tableIndex = _currentPattern->GetFormPattern(0)->IdentifierCodePattern->CodeValue;
		_currentSubject = _info->EvaluationSubjectInfo->EvaluationSubjects->at(_currentPattern->EvaluationSubjectIndex);
		_currentMemberInfo = _info->EvaluationMemberInfo->at(patternCount);
		//直接进行保存
		QString patternName = _currentPattern->FileName;
		QString excelFileName = patternName + ".xlsx";
		//如果文件名重复
		QDir* dir = new QDir(excelName);
		QList<QFileInfo>* fileInfo = new QList<QFileInfo>(dir->entryInfoList());
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

			//初始化表头
			//单位页
			//_mExcelReader->chooseSheet(0);
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
			cellListCount.clear();
			
			_mExcelReader->setSheetName(_currentPattern->MemberTypes);
			for (int i = 0; i < _currentPattern->MemberTypes.count(); i++)
			{
				QList<int> temp;
				cellListCount.append(temp);
				initSheet(i, patternCount);
			}

			
			/*_mExcelReader->chooseSheet(1);
			_mExcelReader->writeExcel(1, 1, u8"成员指标汇总表", format1);
			_mExcelReader->writeExcel(2, 1, u8"单位", format1);
			_mExcelReader->writeExcel(2, 2, u8"姓名", format1);
			_mExcelReader->writeExcel(2, 3, u8"性别", format1);
			_mExcelReader->writeExcel(2, 4, u8"职务", format1);
			_mExcelReader->writeExcel(2, 5, u8"职务类型", format1);
			_mExcelReader->writeExcel(2, 6, u8"表类型", format1);
			_mExcelReader->writeExcel(2, 7, u8"权重", format1);
			_mExcelReader->writeExcel(2, 8, u8"收回数", format1);
			*/
			/*
			int indexCountF = _currentPattern->MemberIndexs->at(patternCount)->MemberDetailIndexs->count();
			int indexCountGroup = 0;
			int indexCountT = 9;
			QList<QString> indexNameF;
			QList<MemberDetailIndex*>* indexNode = new QList<MemberDetailIndex*>;
			indexNameF.clear();
			indexNode->clear();
			QList<int> cellListCount;
			for (unsigned i = 0; i < indexCountF; i++)
			{
				indexNode->append(_currentPattern->MemberIndexs->at(0)->MemberDetailIndexs->at(i));
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
			_mExcelReader->mergeCells(2, 8, 4, 8, format1);


			int unitCount = _currentPattern->EvaluationUnits.count();
			int subjectCount = _currentSubject->EvaluationSubjects->count();
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

				int memberCount = _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->count();
				_mExcelReader->writeExcel(takenSpace, 1, unitName.at(i), format1);
				for (int j = 0; j < memberCount; j++)
				{
					MemberInfo* newMember = _currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j);
					_mExcelReader->writeExcel(takenSpace, 2, newMember->name, format1);
					_mExcelReader->writeExcel(takenSpace, 3, (newMember->gender == 0 ? u8"男" : u8"女"), format1); //int 0male/1female
					_mExcelReader->writeExcel(takenSpace, 4, newMember->duty, format1);
					_mExcelReader->writeExcel(takenSpace, 5, newMember->dutyClass, format1);
					for (int k = 0; k < subjectCount; k++)
					{
						_mExcelReader->writeExcel(takenSpace + k, 6, _currentSubject->EvaluationSubjects->at(k)->name, format1);//!!!!!!!!!!!!
					}
					_mExcelReader->writeExcel(takenSpace + subjectCount, 6, u8"合计", format1);

					_mExcelReader->mergeCells(takenSpace, 2, takenSpace + subjectCount, 2, format1);
					_mExcelReader->mergeCells(takenSpace, 3, takenSpace + subjectCount, 3, format1);
					_mExcelReader->mergeCells(takenSpace, 4, takenSpace + subjectCount, 4, format1);
					_mExcelReader->mergeCells(takenSpace, 5, takenSpace + subjectCount, 5, format1);

					//_currentMemberInfo->EvaluationMembers->at(i)->EvaluationMembers->at(j)->name;
					takenSpace += subjectCount + 1;
				}
				_mExcelReader->mergeCells(takenSpace - memberCount * (subjectCount + 1), 1, takenSpace - 1, 1, format1);
			}
			*/
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
			//vector<vector<vector<vector<int>>>> scoreCount;
			QList<QList<QList<QList<QString>>>> scoreCount;//单位-主体-成员-得分情况
			QList<QList<QList<int>>> memberTypeRecord; 
			QList<int> receiveCount;//收回数 计数subjectIndex
			for (int resultCount = 0; resultCount < results->count(); resultCount++)
			{
				//currentFormPattern = nullptr;
				_currentResult = results->at(resultCount);
				
				if (_currentResult->FormReuslts->at(0)->IdentifierResult->Result.at(0) == tableIndex) //判断模式一致
				{
					QList<MFormResult*>* currentFormResults = new QList<MFormResult*>;
					for (int unitIndex = 0; unitIndex < _currentPattern->EvaluationUnits.count(); unitIndex++)
					{
						QList<QList<QList<QString>>> scoreCount_unit;
						QList<QList<int>> memberTypeRecord_unit;
						//scoreCount.append(scoreCount_unit);
						currentFormResults->clear();
						for (int formIndex = 0; formIndex < _currentResult->FormReuslts->count(); formIndex++)
						{
							if (_currentResult->FormReuslts->at(formIndex)->IsRecognizeSuccess)
							{
								if (_currentResult->FormReuslts->at(formIndex)->IdentifierResult->Result.at(6).toLatin1() - 48 == unitIndex)
								{
									currentFormResults->append(_currentResult->FormReuslts->at(formIndex));
								}
							}
						}
						for (int subjectIndex = 0; subjectIndex < subjectCount; subjectIndex++)//主体组
						{
							receiveCount.append(0);
							QList<QList<QString>> scoreCount_subject;
							QList<int> memberTypeRecord_subject;
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
											memberTypeRecord_subject.append(_currentMemberInfo->EvaluationMembers->at(unitIndex)->EvaluationMembers->at(memberIndex + shift)->MemberType);
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
							memberTypeRecord_unit.append(memberTypeRecord_subject);
						}
						scoreCount.append(scoreCount_unit);
						memberTypeRecord.append(memberTypeRecord_unit);
					}
				}
			}
			
			//QString resultMark = _currentResult->FormReuslts->at(0)->IdentifierResult->Result;
			//int subjectMark = resultMark.at(4).toLatin1() - 48;//主体
			//int unitMark = resultMark.at(6).toLatin1() - 48;//单位

			int pageCount = _currentPattern->GetFormPatternCount();
			receiveCount.append(0);
			for (int i = 0; i < receiveCount.count() - 1; i++)
			{
				receiveCount[i] /= pageCount;
				receiveCount[receiveCount.count() - 1] += receiveCount.at(i);
			}

			QList<int> subjectWeight;
			int subjectWeightSum = 0;
			for (int i = 0 ; i < subjectCount; i ++)
			{
				subjectWeight.append(0);
				subjectWeight[i] = _currentSubject->EvaluationSubjects->at(i)->weightTop;
				subjectWeightSum += subjectWeight.at(i);
			}
			subjectWeight.append(subjectWeightSum);

			//输出
			for (int memberTypeIndex = 0; memberTypeIndex < _currentPattern->MemberTypes.count(); memberTypeIndex++)
			{
				_mExcelReader->chooseSheet(memberTypeIndex);
				int excelRowCount = 5;
				int excelColumnIndex = 9;

				for (int unitIndex = 0; unitIndex < unitCount; unitIndex++)
				{
					int memberCount = _currentMemberInfo->EvaluationMembers->at(unitIndex)->EvaluationMembers->count();
					QList<QList<int>> memberScore;
					for (int subjectIndex = 0; subjectIndex < subjectCount; subjectIndex++)
					{
						if (subjectIndex != 0)
						{
							excelRowCount -= memberCount * (subjectCount + 1);
						}
						//_mExcelReader->writeExcel(excelRowCount + subjectCount, 7, QString::number(ceil(receiveCount[subjectCount] * 1.0 / pageCount * 1.0)), format1);
						for (int memberIndex = 0; memberIndex < memberCount; memberIndex++)
						{
							if (memberTypeRecord.at(unitIndex).at(subjectIndex).at(memberIndex) != memberTypeIndex)
							{
								continue;
							}
							if (subjectIndex == 0)
							{
								QList<int> memberScore_sub;
								memberScore_sub.append(0);
								_mExcelReader->writeExcel(excelRowCount + subjectIndex, 7, QString::number(subjectWeight.at(subjectIndex)), format1);
								_mExcelReader->writeExcel(excelRowCount + subjectIndex, 8, QString::number(receiveCount.at(subjectIndex)), format1);
								for (int i = 0; i < indexCountGroup; i++)
								{
									int empty = cellListCount.at(memberTypeIndex).at(i) - scoreCount[unitIndex][subjectIndex][memberIndex][i].count();
									for (int j = 0; j < empty; j++)
									{
										_mExcelReader->writeExcel(excelRowCount + subjectIndex, excelColumnIndex, "0", format1);
										memberScore_sub[excelColumnIndex - 9] += 0;
										memberScore_sub.append(0);
										excelColumnIndex++;
									}
									for (int n = 0; n < cellListCount.at(memberTypeIndex).at(i) - empty; n++)
									{
										QString result = QString::number((scoreCount[unitIndex][subjectIndex][memberIndex][i][n].toLatin1() - 48) * subjectWeight.at(subjectIndex));
										_mExcelReader->writeExcel(excelRowCount + subjectIndex, excelColumnIndex, result, format1);
										memberScore_sub[excelColumnIndex - 9] += result.toInt();
										memberScore_sub.append(0);
										excelColumnIndex++;
									}
								}
								memberScore.append(memberScore_sub);
								excelColumnIndex = 9;
							}
							else
							{
								_mExcelReader->writeExcel(excelRowCount + subjectIndex, 7, QString::number(subjectWeight.at(subjectIndex)), format1);
								_mExcelReader->writeExcel(excelRowCount + subjectIndex, 8, QString::number(receiveCount.at(subjectIndex)), format1);
								for (int i = 0; i < indexCountGroup; i++)
								{
									//MFormPattern* currentformpattern = _currentPattern->GetFormPattern(0);
									//_currentPattern->MemberIndexs->at(0)->MemberDetailIndexs->at(i)->FirstLevelIndex.groupIndex;
									int empty = cellListCount.at(memberTypeIndex).at(i) - scoreCount[unitIndex][subjectIndex][memberIndex][i].count();
									for (int j = 0; j < empty; j++)
									{
										_mExcelReader->writeExcel(excelRowCount + subjectIndex, excelColumnIndex, "0", format1);
										memberScore[memberIndex][excelColumnIndex - 9] += 0;
										excelColumnIndex++;
									}
									for (int n = 0; n < cellListCount.at(memberTypeIndex).at(i) - empty; n++)
									{
										QString result = QString::number((scoreCount[unitIndex][subjectIndex][memberIndex][i][n].toLatin1() - 48) * subjectWeight.at(subjectIndex));
										_mExcelReader->writeExcel(excelRowCount + subjectIndex, excelColumnIndex, result, format1);
										memberScore[memberIndex][excelColumnIndex - 9] += result.toInt();
										excelColumnIndex++;
									}
								}
								excelColumnIndex = 9;
							}
							if (subjectIndex == subjectCount - 1)
							{
								_mExcelReader->writeExcel(excelRowCount + subjectCount, 7, QString::number(subjectWeight.at(subjectCount)), format1);
								_mExcelReader->writeExcel(excelRowCount + subjectCount, 8, QString::number(receiveCount.at(subjectCount)), format1);
								for (int i = 0; i < memberScore[memberIndex].size() - 1; i++) //总计
								{
									_mExcelReader->writeExcel(excelRowCount + subjectCount, excelColumnIndex, QString::number(memberScore.at(memberIndex).at(i)), format1);
									excelColumnIndex++;
								}
							}
							excelColumnIndex = 9;
							excelRowCount += subjectCount + 1;
						}
					}
				}
			}
		}
	}
}
