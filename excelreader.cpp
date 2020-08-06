#include "excelreader.h"
#include "xlsxdocument.h"

ExcelReader::ExcelReader(QWidget *parent) : QDialog(parent)
{
	_mExcel = nullptr;
	//调用函数
	//readExcel();
    //changeExcel(1, 2, (u8"撒大声地"));
	//testPicture();
	//mergeCells(1,1, 3,3);
}

ExcelReader::~ExcelReader()
{
	if (_mExcel != nullptr)
	{
		delete _mExcel;
	}
	_mExcel = nullptr;
}


/***************************************
*函数功能：新建一个Excel
*输入：
*	void
*输出：
*	bool：是否新建成功
*作者：JZQ
*时间版本：2019-07-03-V1.0
***************************************/
bool ExcelReader::newExcel(QString mexcelName)
{
	if (_mExcel != nullptr)
	{
		//delete _mExcel;
		_mExcel = nullptr;
	}
	_mExcel = new QXlsx::Document(mexcelName);
	if (_mExcel == nullptr)
	{
		return false;
	}

	return true;
}


/***************************************
*函数功能：打开Excel文件
*输入：
*	void
*输出：
*	bool：是否打开成功
*作者：JZQ
*时间版本：2019-07-03-V1.0
***************************************/
bool ExcelReader::openExcel()
{
	//打开Excel文件
	QString excelFileName = QFileDialog::getOpenFileName(0, QString(), QString(), "Excel(*.xlsx)");
	if (excelFileName.isEmpty())
	{
		return false;
	}
	if (_mExcel != nullptr)
	{
		delete _mExcel;
		_mExcel = nullptr;
	}
	_mExcel = new QXlsx::Document(excelFileName);
	if (_mExcel == nullptr)
	{
        QMessageBox::critical(0, (u8"错误信息"), (u8"EXCEL对象丢失"));
		return false;
	}

	return true;
}


/***************************************
*函数功能：读取Excel，只能读取文本。
*输入：
*	void
*输出：
*	bool：是否读取成功
*作者：JZQ
*时间版本：2019-06-12-V2.0
***************************************/
bool ExcelReader::readExcel()
{
	if (_mExcel == nullptr)
	{
        QMessageBox::warning(0, (u8"错误信息"), (u8"请先新建或打开一个EXCEL对象！"));
		return false;
	}

	QXlsx::Workbook *workBook = _mExcel->workbook();
	QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));//打开第一个sheet
	QXlsx::CellRange usedRange = _mExcel->dimension();//使用范围
	QString value;
	for (int i = 1; i <= usedRange.rowCount(); i++)
	{
		for (int j = 1; j <= usedRange.columnCount(); j++)
		{
			QXlsx::Cell *cell = workSheet->cellAt(i, j);
			if (cell == NULL) continue;
			else
			{
				value = cell->value().toString();
				qDebug() << i << " " << j << " " << value;
			}
		}
	}
    qDebug() << (u8"读取成功");
	return true;
}


/***************************************
*函数功能：写入数据到Excel，若有值，则覆盖原来的值。
*输入：
*	iRow:要写入的行
*	iColumn:要写入的列
*	content:要写入的内容
*输出：
*	bool：是否写入成功
*作者：JZQ
*时间版本：2019-07-03-V1.0
***************************************/
bool ExcelReader::writeExcel(const int iRow, const int iColumn, const QString content, QXlsx::Format format)
{
	if (_mExcel == nullptr)
	{
        QMessageBox::warning(0, (u8"错误信息"), (u8"请先新建或打开一个EXCEL对象！"));
		return false;
	}
	if (iRow <= 0 || iColumn <= 0)
	{
        QMessageBox::warning(0, (u8"错误信息"), (u8"写入的行或列索引小于等于0！"));
		return false;
	}
	if (!_mExcel->write(iRow, iColumn, content,format))
	{
        QMessageBox::warning(0, (u8"错误信息"), (u8"写入出错！"));
		return false;
	}
	return _mExcel->save();
}


bool ExcelReader::setSheetName(QList<QString> memberTypeList)
{
	for (int i = 0; i < memberTypeList.count(); i++)
	{
		if (!_mExcel->insertSheet(i, memberTypeList.at(i)))
		{
			_mExcel->insertSheet(i, "abcdefghijklabsdkfj1");
			_mExcel->deleteSheet(memberTypeList.at(i));
			_mExcel->renameSheet("abcdefghijklabsdkfj1", memberTypeList.at(i));
		}
	}
	return true;
}

bool ExcelReader::chooseSheet(int i)
{
	return _mExcel->workbook()->setActiveSheet(i);
}

/***************************************
*函数功能：合并单元格(横向合并或者纵向合并)
*输入：
*	通过第一行，最后一行，第一列，最后一列四个参数来确定合并的区域
*	firstRow:第一行
*	firstColumn:第一列
*	lastRow:最后一行
*	lastColumn:最后一列
*输出：
*	void
*作者：JZQ
*时间版本：2019-06-12-V2.0
***************************************/
void ExcelReader::mergeCells(const int firstRow, const int firstColumn, const int lastRow, const int lastColumn,QXlsx::Format format)
{
	if (_mExcel == nullptr)
	{
        QMessageBox::warning(0, (u8"错误信息"), (u8"请先新建或打开一个EXCEL对象！"));
		return;
	}

	//QXlsx::CellRange usedRange = _mExcel->dimension();//使用范围
	//if (usedRange.rowCount() != 0 || usedRange.columnCount() != 0)
	//{
    //	QMessageBox::critical(0, (u8"错误信息"), (u8"EXCEL对象不为空"));
	//	return;
	//}

	////设置格式：对齐方式、字体大小、颜色、加粗、斜体等。。。。。
	//QXlsx::Format format;
	//format.setHorizontalAlignment(QXlsx::Format::AlignHCenter);//横向居中
	//format.setVerticalAlignment(QXlsx::Format::AlignVCenter);//纵向居中
	//format.setFontColor(QColor(Qt::red));/*文字为红色*/
	//format.setPatternBackgroundColor(QColor(152, 251, 152));/*背景颜色*/
	//format.setFontSize(15);/*设置字体大小*/
	//format.setBorderStyle(QXlsx::Format::BorderDashDotDot);/*边框样式*/
	//format.setFontBold(true);/*设置加粗*/
	//format.setFontUnderline(QXlsx::Format::FontUnderlineSingle);/*下划线(单，还有双。。。。)*/
	//
	//// 设置列宽
	//_mExcel->setColumnWidth(1, 20);
	//_mExcel->setColumnWidth(2, 30);

	QXlsx::CellRange mergeRange(firstRow, firstColumn, lastRow, lastColumn);
	_mExcel->mergeCells(mergeRange, format);
	//_mExcel->write(firstRow, firstColumn, "Hello Qt!");

	if (_mExcel->save())
	{
        qDebug() << (u8"合并成功");
	}
}


/***************************************
*函数功能：拆分某个区域单元格
*输入：
*	通过第一行，最后一行，第一列，最后一列四个参数来确定拆分的区域
*	firstRow:第一行
*	firstColumn:第一列
*	lastRow:最后一行
*	lastColumn:最后一列
*输出：
*	void
*作者：JZQ
*时间版本：2019-06-12-V2.0
***************************************/
void ExcelReader::unmergeCells(const int firstRow, const int firstColumn, const int lastRow, const int lastColumn)
{
	if (_mExcel == nullptr)
	{
        QMessageBox::warning(0, (u8"错误信息"), (u8"请先新建或打开一个EXCEL对象！"));
		return;
	}

	QXlsx::CellRange mergeRange(firstRow, firstColumn, lastRow, lastColumn);

	if (_mExcel->unmergeCells(mergeRange))
	{//取消合并（拆分一个单元格内）
        qDebug() << (u8"取消合并成功");
	}
	else
	{
        qDebug() << (u8"取消合并不成功");
	}
}
