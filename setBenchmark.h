#pragma once

#include <QDialog>
#include "ui_setBenchmark.h"

#include "mevaluationinfo.h"
#include "tableitem.h"

using namespace Cm3::CommonUtils;
using namespace Cm3::FormPattern;

class SetBenchmark : public QDialog
{
	Q_OBJECT

public:
	SetBenchmark(QWidget* parent = Q_NULLPTR);
	~SetBenchmark();
public slots:
	void finishSetAnswer();
	void setPatternSheet(MRecognizeFormPattern*);

signals:
	void sendMessage(QList<QString>);

private:
	Ui::SetBenchmark ui;
private:
	QString _filePath;
	int _dataOperateType;
	TableItem* _tableview;
	int _columnCount;
	QTableView* tableview;
	TableItem* table;
	QList<QString> benchmark;
};
