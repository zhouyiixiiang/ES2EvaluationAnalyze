#pragma once

#include <QDialog>
#include "ui_waitingwindow.h"
#include "mresult.h"
using namespace Cm3::FormResult;
class WaitingWindow : public QDialog
{
	Q_OBJECT

public:
	WaitingWindow( QWidget *parent = Q_NULLPTR);
	~WaitingWindow();
public:
	void setReadData(MResult* result,QString FilePath);
	void setWriteData(MResult* result, QString FilePath);

private:
	Ui::WaitingWindow ui;
private:
	MResult* _result=nullptr;
	QString _filePath;
	int _dataOperateType;
	void showReadDataWindow();
	void showWriteDataWindow();
};
