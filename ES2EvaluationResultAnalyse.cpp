#include "ES2EvaluationResultAnalyse.h"

ES2EvaluationResultAnalyse::ES2EvaluationResultAnalyse(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

    setWindowTitle((u8"数据汇总统计"));

	initGender();

    McurrentResults = new  QList<MResult*>;

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
    ui.tableView_2->setMouseTracking(true);

}

ES2EvaluationResultAnalyse::~ES2EvaluationResultAnalyse()
{

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


void ES2EvaluationResultAnalyse::startLoadThread()
{
    loadThread = new QThread;
    loadDataFile = new LoadDataFile();
    loadDataFile->moveToThread(loadThread);
    connect(loadThread, SIGNAL(finished()), loadDataFile, SLOT(deleteLater()));

	connect(this, &ES2EvaluationResultAnalyse::dataOperate, loadDataFile, &LoadDataFile::doDataOperate);
	connect(loadDataFile, &LoadDataFile::countStep, this, &ES2EvaluationResultAnalyse::uploadCountProgress);
	connect(loadDataFile, &LoadDataFile::finish, this, &ES2EvaluationResultAnalyse::finishLoadDataFile);

	//connect(loadThread, &LoadDataFile::finishExcel, this, &ES2EvaluationResultAnalyse::closeExcelThread);
	//
	//开启线程
	loadThread->start();
}

void ES2EvaluationResultAnalyse::OnButtonLoadDataFile()
{
    QString fileName = QFileDialog::getOpenFileName(this, (u8"添加数据文件"), QString(), "Data Files(*.crst)");
    if (fileName.isEmpty())
	{
		return;
	}
	QFile file(fileName);

    SetWait(true);
    loadDataFile->readData(McurrentResults, fileName);
    dataOperate();
}

void ES2EvaluationResultAnalyse::OnButtonLoadEvaluationData()
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
    SetWait(true);
    loadDataFile->readData(McurrentResults, fileName);
    dataOperate();
}

void ES2EvaluationResultAnalyse::selectionTableEvaluationChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableEvaluation->MselectionModel->selectedIndexes().count() > 0)
	{
		int selectedRow = _tableEvaluation->MselectionModel->selectedIndexes().front().row();
		if (_evaluationInfos->size() > selectedRow)
		{
			_currentEvaluationInfo = _evaluationInfos->at(selectedRow);
			_evaluationIndex = selectedRow;
			_formPatternIndex = 0;
			//getMemberDutyClass();
			// 更新tablePatternInfo表格
			fillingTheTableView("tablePatternInfo");
			fillingTheTableView("tableEvaluationMembers");
			fillingTheTableView("tableEvaluationTable");
		}
	}
}

void ES2EvaluationResultAnalyse::selectionTableEvaluationObjectChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableEvaluationObject->MselectionModel->selectedIndexes().size() > 0)
	{
		int selectedRow = _tableEvaluationObject->MselectionModel->selectedIndexes().front().row();
		if (_currentrecognizePattern->EvaluationUnits.size() > selectedRow)
		{
			_evaluationUnitIndex = selectedRow;
		}
	}
	fillingTheTableView("tableEvaluationMembers");
}

void ES2EvaluationResultAnalyse::selectionTableEvaluationMembersChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableEvaluationMembers->MselectionModel->selectedIndexes().size() > 0)
	{
		int selectedRow = _tableEvaluationMembers->MselectionModel->selectedIndexes().front().row();
		if (_currentEvaluationInfo->EvaluationMemberInfo->at(_evaluationUnitIndex)->EvaluationMembers->at(_formPatternIndex)->EvaluationMembers->size() > selectedRow) {
			_currentEvaluationMemberInfo = _currentEvaluationInfo->EvaluationMemberInfo->at(_evaluationUnitIndex)->EvaluationMembers->at(_formPatternIndex)->EvaluationMembers->at(selectedRow);
			_evaluationMemberIndex = selectedRow;
			//记录下多选项下标
			_selectedMembersIndex.clear();
			for (int i = 0; i < _tableEvaluationMembers->MselectionModel->selectedIndexes().size(); i++)
			{
				_selectedMembersIndex.append(_tableEvaluationMembers->MselectionModel->selectedIndexes().at(i).row());
			}
		}
		_tableEvaluationMembers->MselectionModel->selectedIndexes().clear();
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

	//fillingTheTableView("tableEvaluationObject");
	fillingTheTableView("tableEvaluationMembers");
	fillingTheTableView("tableEvaluation");
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
	else if (s == "tablePatternInfo")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTablePatternInfoChanged(QItemSelection, QItemSelection)));
		QObject::disconnect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTablePatternInfoChanged(QStandardItem*)));
	}
	else if (s == "tableEvaluationObject")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationObjectChanged(QItemSelection, QItemSelection)));
		QObject::disconnect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationObjectChanged(QStandardItem*)));
	}
	else if (s == "tableEvaluationMembers")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationMembersChanged(QItemSelection, QItemSelection)));
		QObject::disconnect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationMembersChanged(QStandardItem*)));
	}
	else if (s == "tableEvaluationTable")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationTableChanged(QItemSelection, QItemSelection)));
		QObject::disconnect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationTableChanged(QStandardItem*)));
	}
}

void ES2EvaluationResultAnalyse::updateAllTableviews()
{
	fillingTheTableView("tableEvaluationMembers");
	fillingTheTableView("tablePatternInfo");
	fillingTheTableView("tableEvaluation");
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
	else if (s == "tableEvaluationObject")
	{
		QObject::connect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationObjectChanged(QItemSelection, QItemSelection)));
		QObject::connect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationObjectChanged(QStandardItem*)));
	}
	else if (s == "tableEvaluationMembers")
	{
		QObject::connect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationMembersChanged(QItemSelection, QItemSelection)));
		QObject::connect(tableItem->Mmodel, SIGNAL(itemChanged(QStandardItem*)), this, SLOT(itemTableEvaluationMembersChanged(QStandardItem*)));
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
	if (sTable == "tableEvaluationMembers")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView_2, _tableEvaluationMembers);
	}
	else if (sTable == "tableEvaluation")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView, _tableEvaluation);
	}
	else if (sTable == "tableEvaluationObject")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView_3, _tableEvaluationObject);
	}
	if (table->Mmodel->rowCount() > 0)
	{
		int j = table->Mmodel->rowCount();
		table->Mmodel->removeRows(0, table->Mmodel->rowCount());
	}
	table->MselectionModel->clear();
	if (sTable == "tableEvaluationMembers")
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
	else if (sTable == "tableEvaluation")
	{
		//表格总体设计
		table->Mmodel->setColumnCount(1);
		tableview->verticalHeader()->hide();//隐藏行号 
		table->Mmodel->setHeaderData(0, Qt::Horizontal, (u8"已有测评"));
		tableview->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
		//往表格中填充模数据
		int row = 0;
		int column = 0;
		if (_evaluationInfos != nullptr)
		{
			for (int i = 0; i < _evaluationInfos->size(); i++)
			{
				table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
				MEvaluationInfo* tempSInfo = _evaluationInfos->at(i);
				fillTableCell(tempSInfo->RecognizePatternInfo->Name, table, i, 0);
			}
		}
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

