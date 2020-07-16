#include "ES2EvaluationResultAnalyse.h"

ES2EvaluationResultAnalyse::ES2EvaluationResultAnalyse(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

    setWindowTitle((u8"数据汇总统计"));

	initGender();
	_path = QCoreApplication::applicationDirPath().toUtf8();
	_selectedMembersIndex.clear();
	_connections.clear();

    McurrentResults = new QList<MResult*>;
	_evaluationInfos = new QList<MEvaluationInfo*>; 

	_currentrecognizePattern = new MRecognizeFormPattern();
	_currentEvaluationInfo = new MEvaluationInfo();

	_mCurrentPatterns = new QList<MPattern*>;
	//McurrentRecognizedResult = new MResult;

	if (_evaluationInfos->size() > 0)
	{
		_currentEvaluationInfo = _evaluationInfos->at(0);
		_evaluationSubjectInfo = _currentEvaluationInfo->EvaluationSubjectInfo;
		//_memberDutyClass = _currentEvaluationInfo->MemberDutyClass;
		//getMemberDutyClass();
	}
	else
	{
		//初始化并从本地读取测评主体与测评成员信息
		_evaluationSubjectInfo = new ES2EvaluationSubjects();
		_evaluationSubjectInfo->readFromBinaryFile();
		_currentEvaluationMember = new ES2EvaluationMember();
		_currentEvaluationMember->readFromBinaryFile();
	}

	_tableEvaluation = new TableItem();
	_tableEvaluationObject = new TableItem();
	_tableEvaluationMembers = new TableItem();

	getconnectUseMapFunc("tableEvaluation", ui.tableView, _tableEvaluation);
	getconnectUseMapFunc("tableEvaluationObject", ui.tableView_2, _tableEvaluationObject);
	getconnectUseMapFunc("tableEvaluationMembers", ui.tableView_3, _tableEvaluationMembers);

    _waitOperate = new DlgWait(this);
    _waitOperate->hide();

    startLoadThread();
    setWindowState(Qt::WindowActive);
    setWindowState(Qt::WindowMaximized);

    ui.tableView->setMouseTracking(true);
	ui.pushButton_2->setEnabled(false);
	ui.pushButton_4->setEnabled(false);
}

ES2EvaluationResultAnalyse::~ES2EvaluationResultAnalyse()
{

	ReleaseQList(_evaluationInfos);
	ReleaseQList(McurrentResults);

	for (int i = 0; i < _connections.size(); i++)
	{
		disconnect(_connections.at(i));
	}
	if (_evaluationSubjectInfo != nullptr)
	{
		delete _evaluationSubjectInfo;
		_evaluationSubjectInfo = nullptr;
	}
	disconnectUsingMapFunc("tableEvaluation", _tableEvaluation);
	disconnectUsingMapFunc("tableEvaluationObject", _tableEvaluationObject);
	disconnectUsingMapFunc("tableEvaluationMembers", _tableEvaluationMembers);
	deleteTableItem(_tableEvaluation);
	deleteTableItem(_tableEvaluationObject);
	deleteTableItem(_tableEvaluationMembers);

}

void ES2EvaluationResultAnalyse::SetWait(bool flag)
{
    if(flag)
    {
        QString hintText = u8"正在读取文件,处理第0/0个数据...";
        hintText = hintText.toUtf8();
        _waitOperate->startAnimation(hintText);
        _waitOperate->show();
        int x = 0;
    }
    else
    {
        _waitOperate->stopAnimation();
        _waitOperate->hide();
    }
}

void ES2EvaluationResultAnalyse::SetWaitExcel(bool flag)
{
	if (flag)
	{
		//ui.frame_tool->setEnabled(false);
		//ui.widget_pattern->setEnabled(false);
		QString ss = u8"正在导出数据至Excel...";
		ss = ss.toUtf8();
		_waitOperate->startAnimation(ss);
		_waitOperate->show();
	}
	else
	{
		//ui.frame_tool->setEnabled(true);
		//ui.widget_pattern->setEnabled(true);
		_waitOperate->stopAnimation();
		_waitOperate->hide();
	}
}

void ES2EvaluationResultAnalyse::startLoadThread()
{
    loadThread = new QThread;
    loadDataFile = new LoadDataFile();
    loadDataFile->moveToThread(loadThread);
    connect(loadThread, SIGNAL(finished()), loadDataFile, SLOT(deleteLater()));

	connect(this, &ES2EvaluationResultAnalyse::dataOperate, loadDataFile, &LoadDataFile::doDataOperate);
	connect(loadDataFile, &LoadDataFile::countStep, this, &ES2EvaluationResultAnalyse::uploadCountProgress);
	connect(loadDataFile, &LoadDataFile::finish, this, &ES2EvaluationResultAnalyse::finishLoadDataFile);

	connect(loadDataFile, &LoadDataFile::finishExcel, this, &ES2EvaluationResultAnalyse::finishSaveExcel);
	
	//开启线程
	loadThread->start();
}

template <typename T>
void ES2EvaluationResultAnalyse::ReleaseQList(QList<T*>* qlist)
{
	if (qlist != nullptr)
	{
		// 清空_recognizePatterns
		QList<T*>::iterator item = qlist->begin();
		while (item != qlist->end())
		{
			qlist->removeOne(*item);
			T* index = (T*)*item;
			delete index;
			index = NULL;
			item++;
		}
		delete qlist;
		qlist = nullptr;
	}
}

void ES2EvaluationResultAnalyse::onButtonLoadDataFile()
{
    QString fileName = QFileDialog::getOpenFileName(this, (u8"添加数据文件"), QString(), "Data Files(*.crst)");
    if (fileName.isEmpty())
	{
		return;
	}
	QFile file(fileName);

    SetWait(true);
    loadDataFile->readDataFile(McurrentResults, fileName);
    dataOperate();
}

void ES2EvaluationResultAnalyse::onButtonLoadEvaluationData()
{
    QString fileName = QFileDialog::getOpenFileName(this, (u8"添加测评参数"), QString(), "Data Files(*.evaluationinfo)");
    if (fileName.isEmpty())
	{
		return;
	}
	QFile file(fileName);
	QString uid;
	if (file.exists())
	{
		file.open(QIODevice::ReadOnly);
		QDataStream input(&file);
		input >> uid;
		file.close();
	}

	for (int i = 0; i < _evaluationInfos->size(); i++)
	{
		if (_evaluationInfos->at(i)->EvaluationUID == uid)
		{
			//执行窗口
			QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"当前已有所添加的测评信息！"), QMessageBox::Yes);
			msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
			int res = msgBox.exec();
			return;
		}
	}
	//readEvaluationInfoFromDatafile();
	_currentEvaluationInfo->readRecognizePatternsFromBinaryFile(fileName);
	_evaluationInfos->append(_currentEvaluationInfo);
	_currentrecognizePattern = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(_formPatternIndex);

	loadDataFile->setInfoData(_currentEvaluationInfo);
	updateAllTableviews();

	ui.pushButton_2->setEnabled(true);

}

void ES2EvaluationResultAnalyse::onButtonOutputExcel()
{
	QString dirName = QFileDialog::getExistingDirectory(this, u8"保存数据结果到...").toUtf8();
	/*if (!dirName.isEmpty())
	{
		dirName = dirName + "/" +_currentEvaluationInfo->RecognizePatternInfo->Name;
		QDir dir(dirName);
		if (!dir.exists()) //如果不存在这个文件夹，则创建这个文件夹
		{
			dir.mkdir(dirName);
		}
	}
	else
		return;
	*/
	//新建Excel文件

	SetWaitExcel(true);
	loadDataFile->setExcelData(McurrentResult, dirName);
	dataOperate();
	//规定excel的单元格式样式
	//QXlsx::Format format1;
	//format1.setFontColor(QColor(Qt::black));/*文字为红色*/
	//format1.setPatternBackgroundColor(QColor(152, 251, 152));/*背景颜色，使用rgb*/
	//format1.setFontSize(15);
	//format1.setHorizontalAlignment(QXlsx::Format::AlignHCenter);/*横向居中*/
	//format1.setBorderStyle(QXlsx::Format::BorderMedium);/*边框样式*/
	//
	//if (SelectedResult->FormReuslts->count() > 0)
	//{
	//	QString excelName= McurrentFormPattern->FormName;
	//	if (_mExcelReader->newExcel(excelName))
	//	{
	//		_mExcelReader->setSheetName();
	//		/*
	//			填充明细数据
	//		*/
	//		_mExcelReader->chooseSheet(0);
	//		int excelRow = SelectedResult->FormReuslts->count();
	//		//默认一个识别结果里面，都是根据一个模板进行识别的，因此结果的分组数是一样的，取第一个FormResult的分组数
	//		int excelColumn = SelectedResult->FormReuslts->at(0)->MarkGroupResults->count();
	//		//表头设置
	//		unsigned k;
	//		for (unsigned j = 0; j < excelColumn; j++)
	//		{
	//			_mExcelReader->writeExcel(1, j + 1, SelectedResult->FormReuslts->at(0)->MarkGroupResults->at(j)->GroupName, format1);
	//			k = j;
	//		}
	//		_mExcelReader->writeExcel(1, k + 2, SelectedResult->FormReuslts->at(0)->IdentifierResult->Name, format1);
	//		for (unsigned i = 0; i < excelRow; i++)
	//		{
	//			Cm3::FormResult::MFormResult *tempFormResult = SelectedResult->FormReuslts->at(i);
	//			unsigned k;
	//			for (unsigned j = 0; j < excelColumn; j++)
	//			{
	//				_mExcelReader->writeExcel(i + 2, j + 1, tempFormResult->MarkGroupResults->at(j)->TextResult, format1);
	//				k = j;
	//			}
	//			_mExcelReader->writeExcel(i + 2, k + 2, tempFormResult->IdentifierResult->Result, format1);
	//		}
	//		/*
	//			填充统计数据
	//		*/
	//		struct AnalyzeResult //建立一个结构体记录每一个统计对象的计数
	//		{
	//			QString patternName;
	//			int count;
	//		};
	//		QList<AnalyzeResult> patternResult;
	//		QList<QList<AnalyzeResult>> patternResults;
	//		_mExcelReader->chooseSheet(1);
	//		patternResults.clear();
	//		//for (int k = 0; k < excelRow; k++)
	//		{
	//			for (int i = 0; i < McurrentFormPattern->MarkGroupPattern->size(); i++)
	//			{
	//				patternResult.clear();
	//				for (int j = 0; j < McurrentFormPattern->MarkGroupPattern->at(i)->CellList->size(); j++)
	//				{
	//					AnalyzeResult eachPatternResult;
	//					eachPatternResult.patternName = McurrentFormPattern->MarkGroupPattern->at(i)->CellList->at(j)->CellName;
	//					eachPatternResult.count = 0;
	//					patternResult.append(eachPatternResult);
	//				}
	//				//往patternResult里面填充未选和多选两种情况
	//				AnalyzeResult eachPatternResult1;
	//				AnalyzeResult eachPatternResult2;
	//				eachPatternResult1.patternName = QString("*");
	//				eachPatternResult2.patternName = QString("");
	//				eachPatternResult1.count = 0;
	//				eachPatternResult2.count = 0;
	//				patternResult.append(eachPatternResult1);
	//				patternResult.append(eachPatternResult2);
	//				//最后在patternResults中添加当前patternResult
	//				patternResults.append(patternResult);
	//			}
	//		}
	//		//计算统计用参数
	//		for (int i = 0; i < SelectedResult->GetFormResultCount(); i++)
	//		{
	//			for (int j = 0; j < SelectedResult->FormReuslts->at(i)->GetGroupResultCount(); j++)
	//			{
	//				for (int k = 0; k < patternResult.size(); k++)
	//				{
	//					if (SelectedResult->FormReuslts->at(i)->GetGroupResult(j)->TextResult==patternResult.at(k).patternName)
	//					{
	//						AnalyzeResult tempAnalyzeResult;
	//						tempAnalyzeResult = patternResults.at(j).at(k);
	//						tempAnalyzeResult.count += 1;
	//						patternResults[j].replace(k, tempAnalyzeResult);
	//						//patternResult[k].count++;
	//					}
	//				}
	//			}
	//		}
	//		//初始化行与列的索引
	//		int rowIndex = 1;
	//		int columnIndex = 1;
	//		//表头设计，根据得到的patternNames确定每一个group的大小，合并相应长度的单元格
	//		int lenGroup = patternResults.at(0).size();
	//		for (int i = 0; i < patternResults.size(); i++)
	//		{
	//			_mExcelReader->mergeCells(rowIndex, columnIndex, rowIndex, columnIndex + lenGroup - 1, format1);
	//			for (int j = 0; j < patternResults.at(0).size(); j++)
	//			{
	//				_mExcelReader->writeExcel(rowIndex, columnIndex, McurrentFormPattern->MarkGroupPattern->at(i)->GroupName, format1);
	//				QString tempPatternName = patternResults.at(i).at(j).patternName;
	//				if (tempPatternName == QString("*"))
	//				{
	//					_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, (u8"多选"), format1);
	//				}
	//				else if (tempPatternName == QString(""))
	//				{
	//					_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, (u8"未选"), format1);
	//				}
	//				else
	//				{
	//					_mExcelReader->writeExcel(rowIndex + 1, columnIndex + j, patternResults.at(i).at(j).patternName, format1);
	//				}
	//				_mExcelReader->writeExcel(rowIndex + 1 + 1, columnIndex + j, QString::number(patternResults.at(i).at(j).count), format1);
	//			}
	//			columnIndex = columnIndex + lenGroup;
	//		}
	//	}
	//}
}



void ES2EvaluationResultAnalyse::selectionTableEvaluationChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableEvaluation->MselectionModel->selectedIndexes().count() > 0)
	{
		int selectedRow = _tableEvaluation->MselectionModel->selectedIndexes().front().row();
		if (_currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->size() > selectedRow)
		{
			//_currentEvaluationInfo = _evaluationInfos->at(selectedRow);
			_currentrecognizePattern = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(selectedRow);
			_formPatternIndex = selectedRow;
			//getMemberDutyClass();
			// 更新tablePatternInfo表格

			fillingTheTableView("tableEvaluationMembers");
			fillingTheTableView("tableEvaluationObject");
		}
	}
}


void ES2EvaluationResultAnalyse::uploadCountProgress(int rec, int sum)
{
    QString hintText;
    hintText = u8"正在读写数据，处理第" + QString::number(rec + 1) + "/" + QString::number(sum) + u8"个数据";
    hintText = hintText.toUtf8();
    _waitOperate->changeText(hintText);
}

void ES2EvaluationResultAnalyse::finishLoadDataFile()
{
    SetWait(false);
	ui.pushButton_4->setEnabled(true);
	
}

void ES2EvaluationResultAnalyse::finishSaveExcel()
{
	SetWait(false);
	//执行窗口
	QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"表格保存成功！"), QMessageBox::Yes);
	msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
	int res = msgBox.exec();
	
}

void ES2EvaluationResultAnalyse::readEvaluationInfoFromDatafile()
{
	if (_evaluationInfos != nullptr)
	{
		ReleaseQList(_evaluationInfos);
	}
	_evaluationInfos = new QList<MEvaluationInfo*>;
	QList<QString> fileNames = fileNamesEndWith("evaluationinfo");
	for (int i = 0; i < fileNames.size(); i++)
	{
		MEvaluationInfo* tempRecPtrn = new MEvaluationInfo();
		tempRecPtrn->readRecognizePatternsFromBinaryFile(fileNames.at(i));
		_evaluationInfos->append(tempRecPtrn);
	}
}

QStringList ES2EvaluationResultAnalyse::fileNamesEndWith(QString name)
{
	QDir* dir = new QDir(_path + "/data");
	QStringList filters;
	filters << "*." + name;
	QList<QFileInfo>* fileInfo = new QList<QFileInfo>(dir->entryInfoList(filters));
	QList<QString> fileNames;
	for (int i = 0; i < fileInfo->count(); i++)
	{
		fileNames.append(fileInfo->at(i).filePath());
	}
	delete dir;
	dir = nullptr;
	fileInfo->clear();
	delete fileInfo;
	fileInfo = nullptr;
	return fileNames;
}

void ES2EvaluationResultAnalyse::initGender()
{
	_gender.clear();
	_gender.append(u8"男");
	_gender.append(u8"女");
	_gender.append(u8"未设定");
}

void ES2EvaluationResultAnalyse::disconnectUsingMapFunc(QString s, TableItem* tableItem)
{
	if (s == "tableEvaluation")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationChanged(QItemSelection, QItemSelection)));
		QObject::disconnect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationChanged(QStandardItem*)));
	}
}

void ES2EvaluationResultAnalyse::updateAllTableviews()
{
	fillingTheTableView("tableEvaluation");
	fillingTheTableView("tableEvaluationMembers");
	fillingTheTableView("tableEvaluationObject");
}

void ES2EvaluationResultAnalyse::reconnectTableItem(QString sTable, QTableView*& tableview, TableItem*& table, QTableView*& exactTableview, TableItem*& exactTable)
{
	disconnectUsingMapFunc(sTable, exactTable);
	deleteTableItem(exactTable);
	exactTable = nullptr;
	exactTable = new TableItem();
	getconnectUseMapFunc(sTable, exactTableview, exactTable);
	tableview = exactTableview;
	table = exactTable;
}

void ES2EvaluationResultAnalyse::getconnectUseMapFunc(QString s, QTableView* table, TableItem* tableItem)
{
	table->setModel(tableItem->Mmodel);
	tableItem->MselectionModel = table->selectionModel();// 获得视图上的选择项
	if (s == "tableEvaluation")
	{
		QObject::connect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationChanged(QItemSelection, QItemSelection)));
		QObject::connect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationChanged(QStandardItem*)));
	}
}

void ES2EvaluationResultAnalyse::deleteTableItem(TableItem*& tableItem) // 清空tableitem指针的泛型函数
{
	if (tableItem != nullptr)
	{
		delete tableItem;
		tableItem = nullptr;
	}
}

void ES2EvaluationResultAnalyse::fillTableCell(QString value, TableItem* table, int row, int col)
{
	QStandardItem* aTem = new QStandardItem(value);
	table->Mmodel->setItem(row, col, aTem);
}

void ES2EvaluationResultAnalyse::fillingTheTableView(QString sTable)
{
	QTableView* tableview = nullptr;
	TableItem* table = nullptr;
	if (sTable == "tableEvaluation")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView, _tableEvaluation);
	}
	else if (sTable == "tableEvaluationMembers")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView_3, _tableEvaluationMembers);
	}
	else if (sTable == "tableEvaluationObject")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView_2, _tableEvaluationObject);
	}
	if (table->Mmodel->rowCount() > 0)
	{
		int j = table->Mmodel->rowCount();
		table->Mmodel->removeRows(0, table->Mmodel->rowCount());
	}
	table->MselectionModel->clear();
	if (sTable == "tableEvaluation")
	{
		//表格总体设计
		table->Mmodel->setColumnCount(1);
		tableview->verticalHeader()->hide();//隐藏行号 
		table->Mmodel->setHeaderData(0, Qt::Horizontal, (u8"已有模式"));
		tableview->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
		//往表格中填充模数据
		int row = 0;
		int column = 0;
		if (_currentEvaluationInfo != nullptr)
		{
			for (int i = 0; i < _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->count(); i++)
			{
				table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
				MRecognizeFormPattern* tempFormPattern = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(i);
				fillTableCell(tempFormPattern->FileName, table, i, 0);
			}
		}
	}
	else if (sTable == "tableEvaluationMembers")
	{
		//表格总体设计
		table->Mmodel->setColumnCount(5);
		tableview->verticalHeader()->hide();//隐藏行号 
		table->Mmodel->setHeaderData(0, Qt::Horizontal, (u8"姓名"));
		table->Mmodel->setHeaderData(1, Qt::Horizontal, (u8"性别"));
		table->Mmodel->setHeaderData(2, Qt::Horizontal, (u8"职务"));
		table->Mmodel->setHeaderData(3, Qt::Horizontal, (u8"职务类别"));
		table->Mmodel->setHeaderData(4, Qt::Horizontal, (u8"编码"));
		tableview->setColumnWidth(0, 90);
		tableview->setColumnWidth(1, 40);
		tableview->setColumnWidth(2, 85);
		tableview->setColumnWidth(3, 100);
		tableview->setColumnWidth(4, 100);
		//往表格中填充模数据
		int row = 0;
		int column = 0;
		
		if (_currentEvaluationInfo != nullptr)
		{
			//
			if (_currentEvaluationInfo->EvaluationMemberInfo->at(_evaluationUnitIndex)->EvaluationMembers->size() == 0) {
				return;
			}
			//
			if (_currentEvaluationInfo->EvaluationMemberInfo->at(_evaluationUnitIndex)->EvaluationMembers->at(_formPatternIndex)->EvaluationMembers->size() > 0)
			{
				for (int i = 0; i < _currentEvaluationInfo->EvaluationMemberInfo->at(_evaluationUnitIndex)->EvaluationMembers->at(_formPatternIndex)->EvaluationMembers->size(); i++)
				{
					table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
					MemberInfo* tempSInfo = _currentEvaluationInfo->EvaluationMemberInfo->at(_evaluationUnitIndex)->EvaluationMembers->at(_formPatternIndex)->EvaluationMembers->at(i);
					fillTableCell(tempSInfo->name, table, i, 0);
					fillTableCell(_gender.at(tempSInfo->gender), table, i, 1);
					fillTableCell(tempSInfo->duty, table, i, 2);
					fillTableCell(tempSInfo->dutyClass, table, i, 3);
					fillTableCell(QString::number(i), table, i, 4);
				}
			}
		}
		table->MselectionModel->clear();
	}

	else if (sTable == "tableEvaluationObject")
	{
		table->Mmodel->setColumnCount(1);
		tableview->verticalHeader()->hide();//隐藏行号 
		table->Mmodel->setHeaderData(0, Qt::Horizontal, (u8"测评单位"));
		tableview->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
		//往表格中填充模数据
		int row = 0;
		int column = 0;
		if (_currentEvaluationInfo != nullptr)
		{
			for (int i = 0; i < _currentrecognizePattern->EvaluationUnits.size(); i++)
			{
				table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
				fillTableCell(_currentrecognizePattern->EvaluationUnits.at(i), table, i, 0);
			}
		}
	}
	if (table->Mmodel->rowCount() > 0)
	{
		int ijk = table->Mmodel->rowCount();
		tableview->selectRow(0);
	}
}

