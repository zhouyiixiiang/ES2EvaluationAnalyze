#include "setBenchmark.h"

SetBenchmark::SetBenchmark(QWidget* parent) : QDialog(parent)
{
	ui.setupUi(this);
	_filePath = "";
	_dataOperateType = 0;

}

SetBenchmark::~SetBenchmark()
{
}

