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
#include "excelreader.h"
#include "mevaluationinfo.h"

#include "es2evaluationsubject.h"
#include "es2evaluationsubjects.h"
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

    void readDataFile(QList<MResult*>*, QString);
    void setExcelData(MResult*, QString, int);
    void setInfoData(MEvaluationInfo*);
    void setAnswer(QList<QString>);
    void setTemplate(QString);

signals:
    void countStep(int, int);
    void countFile();
    void finish();
    void finishExcel();
    void outputError(QString);

public slots:
    void doDataOperate();
    void getCount(int, int);

private:
    void readResultsFromFile();
    void generateExcelResult(QList<MResult*>*, QString);
    void generateTestResult(QList<MResult*>*, QString);
    void formatSet(QXlsx::Format&);
    void initSheet(int, int);
    void initSheet_test(int, int);


    QString _uid;
    QList<MResult*>* _results = nullptr;
    MResult* _result = nullptr;
    //QStringList _fileNames;
    QString _filePath;
    QString _dataOperateType;
    QString _path;
    QString _excelName;
    ExcelReader* _mExcelReader = nullptr;//excel操作对象
    MEvaluationInfo* _info = nullptr;
    QXlsx::Format format1;
    QXlsx::Format format2;
    MRecognizeFormPattern* _currentPattern;
    ES2EvaluationMembers* _currentMemberInfo;
    ES2EvaluationSubject* _currentSubject;
    MResult* _currentResult;
    QString patternName;

    int _outputType;
    int subjectCount;
    int indexCountGroup;
    int unitCount;
    QList<QList<int>> cellListCount;
    QList<QString> _answerList;
    QString _templateName;
};


