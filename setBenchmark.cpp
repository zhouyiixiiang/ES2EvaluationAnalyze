#include "setBenchmark.h"

SetBenchmark::SetBenchmark(QWidget* parent) : QDialog(parent)
{
	ui.setupUi(this);
	_filePath = "";
	_dataOperateType = 0;
	_tableview = new TableItem();
	tableview = ui.tableView;
	table = _tableview;
	tableview->setModel(table->Mmodel);
	table->MselectionModel = tableview->selectionModel();

}

SetBenchmark::~SetBenchmark()
{
	delete _tableview;
}

void SetBenchmark::finishSetAnswer()
{
	benchmark.clear();
	QModelIndex index;
	for (int i = 0;  i < _columnCount;  i++)
	{
		index = table->Mmodel->index(1, i, QModelIndex());
		benchmark.append(index.data().toString());
		if (index.data().toString() == "") 
		{

			//判断空?
		}
	}
	emit sendMessage(benchmark);
	this->accept();
}

void SetBenchmark::setPatternSheet(MRecognizeFormPattern* _pattern)
{
	//MFormPattern* tempFormPattern = _pattern->GetFormPattern(0);
	//_pattern->GetFormPattern(0)->MarkGroupPattern->at(0);
	//?缓存
	table->Mmodel->setRowCount(2);
	table->Mmodel->setHeaderData(0, Qt::Vertical, u8"题号");
	table->Mmodel->setHeaderData(1, Qt::Vertical, u8"正确结果");
	tableview->horizontalHeader()->hide();

	_columnCount = 0;
	bool cache = false;
	if (benchmark.size() > 0)
	{
		cache = true;
	}
	for (int formCount = 0; formCount < _pattern->GetFormPatternCount(); formCount++)
	{
		MemberIndex* tempMemberIndex = _pattern->MemberIndexs->at(formCount);
		for (int i = 0; i < tempMemberIndex->MemberDetailIndexs->count(); i++)
		{
			//_pattern->MemberIndexs->at(i)->MemberDetailIndexs;
			MemberDetailIndex* tempMemberDetail = tempMemberIndex->MemberDetailIndexs->at(i);
			QString typeName = tempMemberDetail->SecondIndexName;
			for (int j = 0; j < tempMemberDetail->SecondLevelIndex.count(); j++)
			{
				QStandardItem* item = new QStandardItem(typeName + tempMemberDetail->SecondLevelIndex.at(j).name);
				table->Mmodel->insertColumn(_columnCount, QModelIndex());
				table->Mmodel->setItem(0, _columnCount, item);
				if (cache)
				{
					table->Mmodel->setItem(1, _columnCount, item);
				}
				_columnCount++;
			}
		}
	}
}
