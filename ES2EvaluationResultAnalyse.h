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

#include <QTableView>
#include <QTableWidget>

#include "tableitem.h"
#include "dlgwait.h"
#include "loadDataFile.h"
#include "mevaluationinfo.h"

#include "es2evaluationsubject.h"
#include "es2evaluationsubjects.h"


using namespace Cm3::CommonUtils;
using namespace Cm3::FormResult;

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
    void selectionTableEvaluationChanged(const QItemSelection& selected, const QItemSelection& deselected);
    void selectionTableEvaluationObjectChanged(const QItemSelection& selected, const QItemSelection& deselected);
    void selectionTableEvaluationMembersChanged(const QItemSelection& selected, const QItemSelection& deselected);

public:
    void SetWait(bool);
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
    void initGender();
    void readEvaluationInfoFromDatafile();
    QStringList fileNamesEndWith(QString);
 
    int _formPatternIndex = 0;
    QList<QString> _gender;
    QList<MEvaluationInfo*>* _evaluationInfos = nullptr; // 已有测评
    MEvaluationInfo* _currentEvaluationInfo = nullptr; // 当前测评
    MRecognizeFormPattern* _currentrecognizePattern = nullptr;
    MemberInfo* _currentEvaluationMemberInfo = nullptr;
    ES2EvaluationSubjects* _evaluationSubjectInfo = nullptr;
    ES2EvaluationMember* _currentEvaluationMember = nullptr;
    QList<int> _selectedMembersIndex;

    QString _path;
    QList<QMetaObject::Connection> _connections;

    QThread* loadThread;
    LoadDataFile* loadDataFile;
    DlgWait* _waitOperate;

    TableItem* _tableEvaluation = nullptr;
    TableItem* _tableEvaluationObject = nullptr;
    TableItem* _tableEvaluationMembers = nullptr;

    int _evaluationIndex = 0;
    int _evaluationUnitIndex = 0;
    int _evaluationTableIndex = 0;
    int _evaluationMemberIndex = 0;
};
