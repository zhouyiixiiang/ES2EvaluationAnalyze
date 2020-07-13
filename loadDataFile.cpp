#include "loadDataFile.h"
#include <windows.h>

LoadDataFile::LoadDataFile(QObject *parent) : QObject(parent)
{
	_mExcelReader = new ExcelReader();
	_filePath = "";
	_dataOperateType = "";
	_path = QCoreApplication::applicationDirPath().toUtf8();
	_uid = "";

    //_filePath = "";
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

void LoadDataFile::setExcelData(MResult* result, MPattern* pattern, QString excelName)
{
	_result = result;
	_pattern = pattern;
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
		for (int i = 0; i < _results->count(); i++)
		{
			_result = _results->at(i);
			generateExcelResult(_result, _pattern, _excelName);
		}

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
		for (int ii = 0; ii < lenResults; ii++)
		{
			MResult* rst = new MResult;
			_results->append(rst);
			
			//MResult* rst = _results->at(ii);
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

void LoadDataFile::generateExcelResult(MResult* curentResult, MPattern* pattern, QString excelName)
{
	//规定excel的单元格式样式
	QXlsx::Format format1;
	format1.setFontColor(QColor(Qt::black));/*文字为红色*/
	format1.setPatternBackgroundColor(QColor(152, 251, 152));/*背景颜色，使用rgb*/
	format1.setFontSize(15);
	format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
	format1.setBorderStyle(QXlsx::Format::BorderMedium);/*边框样式*/

	//直接进行保存
	for (int ii = 0; ii < pattern->GetFormPatternCount(); ii++)
		//int ii = 0;
	{
		MFormPattern* formPattern = pattern->GetFormPattern(ii);
		MResult* result = new MResult();
		for (int i = 0; i < curentResult->GetFormResultCount(); i++)
		{
			if (curentResult->GetFormResult(i)->FormIndex == ii)
			{
				result->AddFormResult(curentResult->GetFormResult(i));
			}
		}
		if (result->GetFormResultCount() > 0)
		{
			QString excelTempName = excelName + "/" + formPattern->FormName + ".xlsx";
			//创建一个改名字的excel表
			//QFile file(excelTempName);
			//file.open(QIODevice::ReadWrite);
			//file.close();
			if (_mExcelReader->newExcel(excelTempName))
			{
				_mExcelReader->setSheetName();
				/*
					填充明细数据
				*/
				_mExcelReader->chooseSheet(0);
				int excelRow = result->GetFormResultCount();
				//默认一个识别结果里面，都是根据一个模板进行识别的，因此结果的分组数是一样的，取第一个FormResult的分组数
				int excelColumn = result->FormReuslts->at(0)->MarkGroupResults->count();
				//表头设置
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
				/*
					填充统计数据
				*/
				if (formPattern->ResultType == SingleSelect)
				{
					struct AnalyzeResult //建立一个结构体记录每一个统计对象的计数
					{
						QString patternName;
						int count;
					};
					QList<AnalyzeResult> patternResult;
					QList<QList<AnalyzeResult>> patternResults;
					_mExcelReader->chooseSheet(1);
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
