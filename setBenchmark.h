#pragma once

#include <QDialog>
#include "ui_setBenchmark.h"

class SetBenchmark : public QDialog
{
	Q_OBJECT

public:
	SetBenchmark(QWidget* parent = Q_NULLPTR);
	~SetBenchmark();

private:
	Ui::SetBenchmark ui;
private:
	QString _filePath;
	int _dataOperateType;
};
