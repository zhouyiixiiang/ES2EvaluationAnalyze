#pragma once

#include <QDialog>
#include "xlsxdocument.h"
#include "xlsxdocument_p.h"
#include "xlsxformat.h"
#include "xlsxcellformula.h"

#include <QString>
#include <QtCore>
#include <QDebug>
#include <QFileDialog>
#include <QMessageBox>

class ExcelReader : public QDialog
{
	Q_OBJECT

public:
	ExcelReader(QWidget *parent = Q_NULLPTR);
	~ExcelReader();

/***************************************
*定义：成员函数
*属性：公有
***************************************/
public:
	bool newExcel(QString mexcelName);//新建一个不存在的excel
	bool openExcel();//打开一个已存在的excel
	bool readExcel();//读取Excel，只能读取文本。//todo:可能不会使用这个函数，若确定不使用，可以删除。。。
	bool writeExcel(const int iRow, const int iColumn, const QString content, QXlsx::Format format);
	void mergeCells(const int firstRow, const int firstColumn, const int lastRow, const int lastColumn, QXlsx::Format format);//合并单元格(横向合并或者纵向合并)
	void unmergeCells(const int firstRow, const int firstColumn, const int lastRow, const int lastColumn);//拆分某个区域单元格
	bool chooseSheet(int i);
	bool setSheetName(QList<QString>);//选定第一个表格，填充明细数据
	QStringList readLine(int);
	QStringList readFormula();
	int getRowCount(int);
	bool copySheet(int, int);
	int getSheetCount();
	int getSheetIndex(QString);
	bool isCellEmpty(int, int);
	void setColumnWidth(int, int);

/***************************************
*定义：成员变量
*属性：私有
***************************************/
private:
	QXlsx::Document *_mExcel = nullptr;
};
