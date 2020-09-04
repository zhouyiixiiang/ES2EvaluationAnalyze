#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_ES2EvaluationResultAnalyse.h"

#include <iostream>
#include <fstream>

#include <QFileDialog>
#include <QFile>
#include <QMessageBox>
#include <QInputDialog>
#include <QDebug>

#include <QTreeWidget>
#include <QList>
#include <QThread>

#include <QTableView>
#include <QTableWidget>

#include "tableitem.h"
#include "dlgwait.h"
#include "loadDataFile.h"
#include "mevaluationinfo.h"
#include "setBenchmark.h"

#include "es2evaluationsubject.h"
#include "es2evaluationsubjects.h"


using namespace Cm3::CommonUtils;
using namespace Cm3::FormResult;
using namespace Es2::EvaluationSubject;
using namespace Cm3::FormPattern;

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
    void setPattern(MRecognizeFormPattern*, QList<QString>);

public slots:
    void onButtonLoadDataFile();
    void onButtonLoadEvaluationData();
    void onButtonOutputCurrentExcel();
    void onButtonOutputExcel();
    void onButtonAddTemplate();
    void onButtonDeleteTemplate();
    void onButtonReadBenchmark();
    void finishLoadDataFile();
    void finishSaveExcel();
    void finishSetAnswer(QList<QString>);
    void outputErrorPrompt(QString);
    void outputUnexpectedPrompt(QString);
    void uploadCountProgress(int, int);
    void selectionTableEvaluationChanged(const QItemSelection& selected, const QItemSelection& deselected);
    void selectionTableTemplateListChanged(const QItemSelection& selected, const QItemSelection& deselected);
    void selectionTableEvaluationObjectChanged(const QItemSelection& selected, const QItemSelection& deselected);
    //void selectOutputType(int);

public:
    void SetWait(bool);
    void SetWaitExcel(bool);
    void fillingTheTableView(QString);
    void fillTableCell(QString value, TableItem* table, int row, int col);

    void getconnectUseMapFunc(QString s, QTableView* table, TableItem* tableItem);
    void disconnectUsingMapFunc(QString s, TableItem* tableItem);
    void updateAllTableviews();
    void reconnectTableItem(QString sTable, QTableView*& tableview, TableItem*& table, QTableView*& exactTableview, TableItem*& exactTable);
    void deleteTableItem(TableItem*& tableItem);


    MResult* McurrentResult = nullptr;
    QList<MResult*>* McurrentResults = nullptr;


private:
    template <typename T>
    void ReleaseQList(QList<T*>*);
    void initLists();
    bool selectTemplateType();
 

    QList<QString> _gender;
    MEvaluationInfo* _currentEvaluationInfo = nullptr; // 当前测评
    MRecognizeFormPattern* _currentrecognizePattern = nullptr;
    MemberInfo* _currentEvaluationMemberInfo = nullptr;
    ES2EvaluationSubjects* _evaluationSubjectInfo = nullptr;
    ES2EvaluationMember* _currentEvaluationMember = nullptr;
    QList<QList<QString>> _benchmark;
    QList<QString> _currentTemplateList;

    QString _path;

    QThread* loadThread;
    LoadDataFile* loadDataFile;
    DlgWait* _waitOperate;
    SetBenchmark* _setAnswer;

    QStringList templateTypeNameList; 
    QList<int> templateType;

    TableItem* _tableEvaluation = nullptr;
    TableItem* _tableEvaluationObject = nullptr;
    TableItem* _tableEvaluationMembers = nullptr;
    TableItem* _tableTemplateList = nullptr;

    //int _evaluationIndex = 0;
    int _evaluationUnitIndex = 0;
    //int _evaluationTableIndex = 0;
    //int _evaluationMemberIndex = 0;
    
    int _formPatternIndex = 0;
    int _templateIndex = 0;
    int _outputType = 0;
    //int _templateType = 0;
};
