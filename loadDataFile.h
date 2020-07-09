#pragma once

#include <QObject>	
#include <QThread>
#include <QDir>
#include <QtCore/QCoreApplication>
#include <QDebug>
#include <iostream>
#include <fstream>

#include <QFileDialog>
#include <QFile>
#include <QMessageBox>
#include <QDebug>

#include <QTreeWidget>
#include <QList>
#include <QThread>

#include "mresult.h"
#include "mpattern.h"
#include "mformpattern.h"
//#include "excelreader.h"

#include "dlgwait.h"

using namespace Cm3::CommonUtils;
using namespace Cm3::FormResult;
using namespace Cm3::FormPattern;

class LoadDataFile : public QObject
{
    Q_OBJECT

public:
    LoadDataFile(QObject *parent = nullptr);
    ~LoadDataFile();

    void readData(QList<MResult*>*, QString);

signals:
    void countStep(int, int);
    void countFile();
    void finish();

public slots:
    void doDataOperate();
    void getCount(int, int);

private:
    void readResultsFromFile();
    //void generateExcelResult(MResult*, MPattern*, QString);

    QString _uid;
    QList<MResult*>* _results = nullptr;
    MResult* _result = nullptr;
    QStringList _fileNames;
    QString _filePath;
    QString _dataOperateType;
    QString _path;

};


