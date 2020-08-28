#include "ES2EvaluationResultAnalyse.h"

ES2EvaluationResultAnalyse::ES2EvaluationResultAnalyse(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

    setWindowTitle((u8"数据汇总统计"));

	initLists();
	_path = QCoreApplication::applicationDirPath().toUtf8();

    McurrentResults = new QList<MResult*>;

	_currentrecognizePattern = new MRecognizeFormPattern();
	_currentEvaluationInfo = new MEvaluationInfo();

	//McurrentRecognizedResult = new MResult;

	//初始化并从本地读取测评主体与测评成员信息
	_evaluationSubjectInfo = new ES2EvaluationSubjects();
	_evaluationSubjectInfo->readFromBinaryFile();
	_currentEvaluationMember = new ES2EvaluationMember();
	_currentEvaluationMember->readFromBinaryFile();
	

	_tableEvaluation = new TableItem();
	_tableEvaluationObject = new TableItem();
	_tableEvaluationMembers = new TableItem();
	_tableTemplateList = new TableItem();

	getconnectUseMapFunc("tableEvaluation", ui.tableView, _tableEvaluation);
	getconnectUseMapFunc("tableEvaluationObject", ui.tableView_2, _tableEvaluationObject);
	getconnectUseMapFunc("tableEvaluationMembers", ui.tableView_3, _tableEvaluationMembers);
	getconnectUseMapFunc("tableTemplateList", ui.tableView_4, _tableTemplateList);

    _waitOperate = new DlgWait(this);
    _waitOperate->hide();
	

    startLoadThread();
    setWindowState(Qt::WindowActive);
    setWindowState(Qt::WindowMaximized);

    ui.tableView->setMouseTracking(true);
	ui.tableView_2->setMouseTracking(true);
	ui.tableView_4->setMouseTracking(true);

	ui.pushButton_2->setEnabled(false);
	ui.pushButton_4->setEnabled(false);
	ui.pushButton_5->setEnabled(false);
	ui.pushButton_6->setEnabled(false);
	ui.pushButton_7->setEnabled(false);
	ui.pushButton_8->setEnabled(false);
	ui.pushButton_7->setVisible(false);
	
	//fillingTheTableView("tableTemplateList");
}

ES2EvaluationResultAnalyse::~ES2EvaluationResultAnalyse()
{

	ReleaseQList(McurrentResults);

	if (_evaluationSubjectInfo != nullptr)
	{
		delete _evaluationSubjectInfo;
		_evaluationSubjectInfo = nullptr;
	}
	if (_currentEvaluationInfo != nullptr)
	{
		delete _currentEvaluationInfo;
		_currentEvaluationInfo = nullptr;
	}

	disconnectUsingMapFunc("tableEvaluation", _tableEvaluation);
	disconnectUsingMapFunc("tableEvaluationObject", _tableEvaluationObject);
	disconnectUsingMapFunc("tableEvaluationMembers", _tableEvaluationMembers);
	disconnectUsingMapFunc("tableTemplateList", _tableTemplateList);

	deleteTableItem(_tableEvaluation);
	deleteTableItem(_tableEvaluationObject);
	deleteTableItem(_tableEvaluationMembers);
	deleteTableItem(_tableTemplateList);
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
	connect(loadDataFile, &LoadDataFile::outputError, this, &ES2EvaluationResultAnalyse::outputErrorPrompt);
	
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
	QString uid;
	if (file.exists())
	{
		file.open(QIODevice::ReadOnly);
		QDataStream input(&file);
		input >> uid;
		file.close();
	}
	if (_currentEvaluationInfo->EvaluationUID != uid)
	{
		//执行窗口
		QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"数据信息与当前测评不匹配！"), QMessageBox::Yes);
		msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
		int res = msgBox.exec();
		return;
	}

    SetWait(true);
	McurrentResults->clear();
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
	//int temp;
	QString uid;
	if (file.exists())
	{
		file.open(QIODevice::ReadOnly);
		QDataStream input(&file);
		input >> uid;
		file.close();
	}

	if (_currentEvaluationInfo->EvaluationUID == uid)
	{
		//执行窗口
        QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"当前已有所添加的测评信息！是否替换？"), QMessageBox::Yes | QMessageBox::No);
        msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
        msgBox.button(QMessageBox::No)->setText((u8"取消"));
        if(msgBox.exec() == QMessageBox::No)
        {
            return;
        }
	}

	delete _currentEvaluationInfo;
	_currentEvaluationInfo = new MEvaluationInfo();
	_currentEvaluationInfo->readRecognizePatternsFromBinaryFile(fileName);
	/*
	ES2EvaluationSubjects* tempSubjectInfo = _currentEvaluationInfo->EvaluationSubjectInfo;
	int tempsize = tempSubjectInfo->EvaluationSubjects->size();
	ES2EvaluationSubject* tempSub = tempSubjectInfo->EvaluationSubjects->at(0);
	*/
	_currentrecognizePattern = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(_formPatternIndex);

	_outputType = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(0)->TableType;
	if (_outputType == 0)
	{
		//考评
		ui.label_2->setText(u8"当前输出类型: 考评");
		ui.pushButton_7->setVisible(true);
		_benchmark.clear();
		for (int i = 0; i < _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->count(); i++)
		{
			QList<QString> temp;
			_benchmark.append(temp);
		}
	}
	else if (_outputType == 1)
	{
		//测评
		ui.label_2->setText(u8"当前输出类型: 测评");
		ui.pushButton_7->setVisible(false);
	}
	else
	{
		ui.label_2->setText(u8"当前输出类型: 未知");
	}

	ui.label_3->setText(u8"当前测评名称: " + _currentEvaluationInfo->RecognizePatternInfo->Name);

	loadDataFile->setInfoData(_currentEvaluationInfo);
	_currentTemplateList.clear();
	templateType.clear();
	_currentTemplateList.append("");
	templateType.append(0);
	updateAllTableviews();

	ui.pushButton_2->setEnabled(true);
	ui.pushButton_4->setEnabled(false);
	ui.pushButton_5->setEnabled(false);
	ui.pushButton_6->setEnabled(false);
	ui.pushButton_7->setEnabled(false);
	ui.pushButton_8->setEnabled(false);
}

void ES2EvaluationResultAnalyse::onButtonOutputCurrentExcel()
{
	if (_currentrecognizePattern->MemberIndexs->isEmpty() || _currentrecognizePattern->MemberIndexs->at(0)->MemberDetailIndexs->isEmpty())
	{
		QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"当前模式未录入指标!"), QMessageBox::Yes | QMessageBox::No);
		msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
		msgBox.button(QMessageBox::No)->setText((u8"取消"));
		if (msgBox.exec() == QMessageBox::No) {
			return;
		}
	}
	if (_outputType == 0 && _benchmark.at(_formPatternIndex).isEmpty())
	{
		QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"还没有录入正确结果！"), QMessageBox::Yes);
		msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
		int res = msgBox.exec();
		return;
	}
	QString dirName = QFileDialog::getExistingDirectory(this, u8"保存数据结果到...").toUtf8();
	if (dirName.isEmpty())
	{
		return;
	}

	SetWaitExcel(true);

	loadDataFile->setExcelData(McurrentResult, dirName, _outputType, _formPatternIndex);
	loadDataFile->setTemplate(_currentTemplateList.at(_templateIndex), templateType.at(_templateIndex));
	dataOperate();
}


void ES2EvaluationResultAnalyse::onButtonOutputExcel()
{
	/*
	if (_currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(0)->TableType != _outputType)//判断类型是否一致:考评/测评
	{
		QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"结果输出类别与数据类别不匹配！"), QMessageBox::Yes);
		msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
		int res = msgBox.exec();
		return;
	}
	*/
	for (int i = 0; i < _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->count(); i++)
	{
		if (_currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(i)->MemberIndexs->isEmpty() || _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(i)->MemberIndexs->at(0)->MemberDetailIndexs->isEmpty())
		{
			QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"第" + QString::number(i + 1) + u8"个模式未录入指标!"), QMessageBox::Yes | QMessageBox::No);
			msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
			msgBox.button(QMessageBox::No)->setText((u8"取消"));
			if (msgBox.exec() == QMessageBox::No) {
				return;
			}
		}
	}
	if (_outputType == 0)
	{
		for (int i = 0; i < _benchmark.size(); i++)
		{
			if (_benchmark.at(i).isEmpty())
			{
				QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"还没有录入正确结果！"), QMessageBox::Yes);
				msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
				int res = msgBox.exec();
				return;
			}
		}
	}
	QString dirName = QFileDialog::getExistingDirectory(this, u8"保存数据结果到...").toUtf8();
	if (!dirName.isEmpty())
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
	
	//新建Excel文件

	SetWaitExcel(true);

	loadDataFile->setExcelData(McurrentResult, dirName, _outputType, -1);
	loadDataFile->setTemplate(_currentTemplateList.at(_templateIndex), templateType.at(_templateIndex));
	dataOperate();

}

bool ES2EvaluationResultAnalyse::selectTemplateType()
{
	QString dlgTitle = u8"模板类型";
	QString txtLabel = u8"选择结果模板类型";
	int     curIndex = 0; //初始选择项
	bool    editable = false; //ComboBox是否可编辑
	bool    finish = false;
	QString text = QInputDialog::getItem(this, dlgTitle, txtLabel, templateTypeNameList, curIndex, editable, &finish);
	if (finish && !text.isEmpty())
	{
		for (int i = 0; i < templateTypeNameList.count(); i++)
		{
			if (text.compare(templateTypeNameList.at(i)) == 0)
			{
				templateType.append(i);
				return true;
			}
		}
		return false;
	}
	else
	{
		return false;
	}
}

void ES2EvaluationResultAnalyse::onButtonAddTemplate()
{
	if (!selectTemplateType())
	{
		return;
	}
	QString fileName = QFileDialog::getOpenFileName(this, (u8"添加Excel模板"), QString(), "Data Files(*.xlsx)");
	if (fileName.isEmpty())
	{
		templateType.removeLast();
		return;
	}

	_currentTemplateList.append(fileName);
	//templateType.append(templateType.at(_templateIndex));


	fillingTheTableView("tableTemplateList");
}

void ES2EvaluationResultAnalyse::onButtonDeleteTemplate()
{
	if (_templateIndex == 0)
	{
		QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"不能删除默认模板！"), QMessageBox::Yes);
		msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
		int res = msgBox.exec();
		return;
	}
	_currentTemplateList.removeAt(_templateIndex);
	templateType.removeAt(_templateIndex);

	fillingTheTableView("tableTemplateList");
}

void ES2EvaluationResultAnalyse::onButtonReadBenchmark()
{
	_setAnswer = new SetBenchmark(this);

	connect(this, &ES2EvaluationResultAnalyse::setPattern, _setAnswer, &SetBenchmark::setPatternSheet);
	connect(_setAnswer, &SetBenchmark::sendMessage, this, &ES2EvaluationResultAnalyse::finishSetAnswer);

	setPattern(_currentrecognizePattern, _benchmark.at(_formPatternIndex));

	if (_currentrecognizePattern->MemberIndexs->isEmpty() || _currentrecognizePattern->MemberIndexs->at(0)->MemberDetailIndexs->isEmpty())
	{
		QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"当前模式未录入指标!"), QMessageBox::Yes);
		msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
		int res = msgBox.exec();
		return;
	}

	_setAnswer->show();
}

void ES2EvaluationResultAnalyse::selectionTableTemplateListChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableTemplateList->MselectionModel->selectedIndexes().count() > 0)
	{
		int selectedRow = _tableTemplateList->MselectionModel->selectedIndexes().front().row();
		_templateIndex = selectedRow;

		//int i = 0;
		//fillingTheTableView("tableTemplateList");
	}
}

void ES2EvaluationResultAnalyse::selectionTableEvaluationChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableEvaluation->MselectionModel->selectedIndexes().count() > 0)
	{
		int selectedRow = _tableEvaluation->MselectionModel->selectedIndexes().front().row();
		if (_currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->size() > selectedRow)
		{
			_currentrecognizePattern = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(selectedRow);
			_formPatternIndex = selectedRow;
			// 更新表格

			fillingTheTableView("tableEvaluationObject");
			fillingTheTableView("tableEvaluationMembers");
		}
	}
}

void ES2EvaluationResultAnalyse::selectionTableEvaluationObjectChanged(const QItemSelection& selected, const QItemSelection& deselected)
{
	if (_tableEvaluationObject->MselectionModel->selectedIndexes().count() > 0)
	{
		int selectedRow = _tableEvaluationObject->MselectionModel->selectedIndexes().front().row();
		if (_currentrecognizePattern->EvaluationUnits.count() > selectedRow)
		{
			_evaluationUnitIndex = selectedRow;
			// 更新表格

			fillingTheTableView("tableEvaluationMembers");
		}
	}
}

/*
void ES2EvaluationResultAnalyse::selectOutputType(int index)
{
	//_outputType = index;
	if (index == 0)
	{//测评
		ui.pushButton_7->setVisible(false);
		_outputType = 1;
	}
	else if (index == 1)
	{//考评
		ui.pushButton_7->setVisible(true);
		_outputType = 0;
	}
}
*/

void ES2EvaluationResultAnalyse::finishSetAnswer(QList<QString> benchmark)
{
	_benchmark[_formPatternIndex] = benchmark;
	loadDataFile->setAnswer(_benchmark);

	fillingTheTableView("tableEvaluation");
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
	ui.pushButton_5->setEnabled(true);
	ui.pushButton_6->setEnabled(true);
	ui.pushButton_7->setEnabled(true);
	ui.pushButton_8->setEnabled(true);
}

void ES2EvaluationResultAnalyse::finishSaveExcel()
{
	SetWait(false);
	//执行窗口
	QMessageBox msgBox(QMessageBox::Information, (u8"提示"), (u8"表格保存成功！"), QMessageBox::Yes);
	msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
	int res = msgBox.exec();
	
}

void ES2EvaluationResultAnalyse::outputErrorPrompt(QString hint)
{
	SetWait(false);
	//执行窗口
	QMessageBox msgBox(QMessageBox::Information, (u8"生成结果错误"), hint, QMessageBox::Yes);
	msgBox.button(QMessageBox::Yes)->setText((u8"确定"));
	int res = msgBox.exec();

}

void ES2EvaluationResultAnalyse::initLists()
{
	_gender.clear();
	_gender.append(u8"男");
	_gender.append(u8"女");
	_gender.append(u8"未设定");

	templateTypeNameList << u8"默认不复制" << u8"按行复制" << u8"按页复制";
}

void ES2EvaluationResultAnalyse::disconnectUsingMapFunc(QString s, TableItem* tableItem)
{
	if (s == "tableEvaluation")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationChanged(QItemSelection, QItemSelection)));
	}
	else if (s == "tableEvaluationObject")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationObjectChanged(QItemSelection, QItemSelection)));
	}
	else if (s == "tableTemplateList")
	{
		QObject::disconnect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectiontableTemplateListChanged(QItemSelection, QItemSelection)));
	}
}

void ES2EvaluationResultAnalyse::updateAllTableviews()
{
	fillingTheTableView("tableEvaluation");
	fillingTheTableView("tableEvaluationMembers");
	fillingTheTableView("tableEvaluationObject");
	fillingTheTableView("tableTemplateList");
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
		QObject::connect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationChanged(QItemSelection, QItemSelection)));\
	}
	else if (s == "tableEvaluationObject")
	{
		QObject::connect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableEvaluationObjectChanged(QItemSelection, QItemSelection)));
	}
	else if (s == "tableTemplateList")
	{
		QObject::connect(tableItem->MselectionModel, SIGNAL(selectionChanged(QItemSelection, QItemSelection)), this, SLOT(selectionTableTemplateListChanged(QItemSelection, QItemSelection)));
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
	QTableView* tableview;
	TableItem* table;
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
	else if (sTable == "tableTemplateList")
	{
		reconnectTableItem(sTable, tableview, table, ui.tableView_4, _tableTemplateList);
	}
	else
	{
		return;
	}
	if (table->Mmodel->rowCount() > 0)
	{
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
		tableview->setEditTriggers(QAbstractItemView::NoEditTriggers);

		//往表格中填充模数据
		int row = 0;
		int column = 0;
		if (_currentEvaluationInfo != nullptr)
		{
			for (int i = 0; i < _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->count(); i++)
			{
				table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
				MRecognizeFormPattern* tempFormPattern = _currentEvaluationInfo->RecognizePatternInfo->RecognizeFormPatterns->at(i);
				QString status = "";
				if (_outputType == 0 && _benchmark.at(i).isEmpty())
				{
					status.append(u8" (未录入答案)");
				}
				fillTableCell(tempFormPattern->FileName + status, table, i, 0);
			}
		}
		_formPatternIndex = 0;
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
		tableview->setEditTriggers(QAbstractItemView::NoEditTriggers);

		//往表格中填充模数据
		int row = 0;
		int column = 0;
		
		if (_currentEvaluationInfo != nullptr)
		{
			//
			if (_currentEvaluationInfo->EvaluationMemberInfo->at(_formPatternIndex)->EvaluationMembers->size() == 0) {
				return;
			}
			//
			if (_currentEvaluationInfo->EvaluationMemberInfo->at(_formPatternIndex)->EvaluationMembers->at(_evaluationUnitIndex)->EvaluationMembers->size() > 0)
			{
				for (int i = 0; i < _currentEvaluationInfo->EvaluationMemberInfo->at(_formPatternIndex)->EvaluationMembers->at(_evaluationUnitIndex)->EvaluationMembers->size(); i++)
				{
					table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
					MemberInfo* tempSInfo = _currentEvaluationInfo->EvaluationMemberInfo->at(_formPatternIndex)->EvaluationMembers->at(_evaluationUnitIndex)->EvaluationMembers->at(i);
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
		table->Mmodel->setHeaderData(0, Qt::Horizontal, (u8"测评单位"));
		tableview->verticalHeader()->hide();//隐藏行号 
		tableview->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
		tableview->setEditTriggers(QAbstractItemView::NoEditTriggers);
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
		_evaluationUnitIndex = 0;
	}
	else if (sTable == "tableTemplateList")
	{
		table->Mmodel->setColumnCount(1);
		table->Mmodel->setHeaderData(0, Qt::Horizontal, (u8"可选模板"));
		tableview->verticalHeader()->hide();//隐藏行号
		tableview->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
		tableview->setEditTriggers(QAbstractItemView::NoEditTriggers);

		table->Mmodel->insertRows(0, 1, QModelIndex());//插入每一行
		fillTableCell(u8"默认Excel文件", table, 0, 0);
		for (int i = 1; i < _currentTemplateList.size(); i++)
		{
			QFileInfo fileInfo(_currentTemplateList.at(i));
			table->Mmodel->insertRows(i, 1, QModelIndex());//插入每一行
			QString hint = "";
			if (templateType.at(i) > 0)
			{
				hint.append(" (" + templateTypeNameList.at(templateType.at(i)) + ")");
			}
			fillTableCell(fileInfo.fileName() + hint, table, i, 0);
		}
		_templateIndex = 0;
	}
	if (table->Mmodel->rowCount() > 0)
	{
		//int ijk = table->Mmodel->rowCount();
		tableview->selectRow(0);
	}
}

