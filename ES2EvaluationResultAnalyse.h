#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_ES2EvaluationResultAnalyse.h"

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
#include "loadDataFile.h"

using namespace Cm3::CommonUtils;

class ES2EvaluationResultAnalyse : public QMainWindow
{
    Q_OBJECT

public:
    ES2EvaluationResultAnalyse(QWidget *parent = Q_NULLPTR);
    ~ES2EvaluationResultAnalyse();

private:
    Ui::ES2EvaluationResultAnalyseClass ui;

private:
    void startLoadThread();

signals:
    void dataOperate();

public slots:
    void OnButtonLoadDataFile();
    void OnButtonLoadEvaluationData();
    void finishLoadDataFile();
    void uploadCountProgress(int, int);

public:
    void SetWait(bool);

private:
    QThread* dataOperateThread;
    QThread* loadThread;
    LoadDataFile* loadDataFile;

    DlgWait* _waitOperate;
    QList<MResult*>* McurrentResults = nullptr;

    


};
