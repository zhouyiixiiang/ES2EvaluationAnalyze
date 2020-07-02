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

#include "dlgwait.h"

using namespace Cm3::CommonUtils;

class LoadDataFile : public QObject
{
    Q_OBJECT

public:
    LoadDataFile(QObject *parent = nullptr);
    ~LoadDataFile();

    void readData();

signals:
    void countStep(int, int);
    void countFile();
    void finish();

public slots:
    void doDataOperate();
    void getCount(int, int);

private:
    QStringList _fileNames;

};


