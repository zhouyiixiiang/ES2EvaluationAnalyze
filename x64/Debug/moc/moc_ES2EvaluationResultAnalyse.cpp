/****************************************************************************
** Meta object code from reading C++ file 'ES2EvaluationResultAnalyse.h'
**
** Created by: The Qt Meta Object Compiler version 67 (Qt 5.12.1)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include "../../../ES2EvaluationResultAnalyse.h"
#include <QtCore/qbytearray.h>
#include <QtCore/qmetatype.h>
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'ES2EvaluationResultAnalyse.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 67
#error "This file was generated using the moc from 5.12.1. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
QT_WARNING_PUSH
QT_WARNING_DISABLE_DEPRECATED
struct qt_meta_stringdata_ES2EvaluationResultAnalyse_t {
    QByteArrayData data[6];
    char stringdata0[100];
};
#define QT_MOC_LITERAL(idx, ofs, len) \
    Q_STATIC_BYTE_ARRAY_DATA_HEADER_INITIALIZER_WITH_OFFSET(len, \
    qptrdiff(offsetof(qt_meta_stringdata_ES2EvaluationResultAnalyse_t, stringdata0) + ofs \
        - idx * sizeof(QByteArrayData)) \
    )
static const qt_meta_stringdata_ES2EvaluationResultAnalyse_t qt_meta_stringdata_ES2EvaluationResultAnalyse = {
    {
QT_MOC_LITERAL(0, 0, 26), // "ES2EvaluationResultAnalyse"
QT_MOC_LITERAL(1, 27, 11), // "dataOperate"
QT_MOC_LITERAL(2, 39, 0), // ""
QT_MOC_LITERAL(3, 40, 20), // "OnButtonLoadDataFile"
QT_MOC_LITERAL(4, 61, 18), // "finishLoadDataFile"
QT_MOC_LITERAL(5, 80, 19) // "uploadCountProgress"

    },
    "ES2EvaluationResultAnalyse\0dataOperate\0"
    "\0OnButtonLoadDataFile\0finishLoadDataFile\0"
    "uploadCountProgress"
};
#undef QT_MOC_LITERAL

static const uint qt_meta_data_ES2EvaluationResultAnalyse[] = {

 // content:
       8,       // revision
       0,       // classname
       0,    0, // classinfo
       4,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       1,       // signalCount

 // signals: name, argc, parameters, tag, flags
       1,    0,   34,    2, 0x06 /* Public */,

 // slots: name, argc, parameters, tag, flags
       3,    0,   35,    2, 0x0a /* Public */,
       4,    0,   36,    2, 0x0a /* Public */,
       5,    2,   37,    2, 0x0a /* Public */,

 // signals: parameters
    QMetaType::Void,

 // slots: parameters
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void, QMetaType::Int, QMetaType::Int,    2,    2,

       0        // eod
};

void ES2EvaluationResultAnalyse::qt_static_metacall(QObject *_o, QMetaObject::Call _c, int _id, void **_a)
{
    if (_c == QMetaObject::InvokeMetaMethod) {
        auto *_t = static_cast<ES2EvaluationResultAnalyse *>(_o);
        Q_UNUSED(_t)
        switch (_id) {
        case 0: _t->dataOperate(); break;
        case 1: _t->OnButtonLoadDataFile(); break;
        case 2: _t->finishLoadDataFile(); break;
        case 3: _t->uploadCountProgress((*reinterpret_cast< int(*)>(_a[1])),(*reinterpret_cast< int(*)>(_a[2]))); break;
        default: ;
        }
    } else if (_c == QMetaObject::IndexOfMethod) {
        int *result = reinterpret_cast<int *>(_a[0]);
        {
            using _t = void (ES2EvaluationResultAnalyse::*)();
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&ES2EvaluationResultAnalyse::dataOperate)) {
                *result = 0;
                return;
            }
        }
    }
}

QT_INIT_METAOBJECT const QMetaObject ES2EvaluationResultAnalyse::staticMetaObject = { {
    &QMainWindow::staticMetaObject,
    qt_meta_stringdata_ES2EvaluationResultAnalyse.data,
    qt_meta_data_ES2EvaluationResultAnalyse,
    qt_static_metacall,
    nullptr,
    nullptr
} };


const QMetaObject *ES2EvaluationResultAnalyse::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->dynamicMetaObject() : &staticMetaObject;
}

void *ES2EvaluationResultAnalyse::qt_metacast(const char *_clname)
{
    if (!_clname) return nullptr;
    if (!strcmp(_clname, qt_meta_stringdata_ES2EvaluationResultAnalyse.stringdata0))
        return static_cast<void*>(this);
    return QMainWindow::qt_metacast(_clname);
}

int ES2EvaluationResultAnalyse::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = QMainWindow::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        if (_id < 4)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 4;
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        if (_id < 4)
            *reinterpret_cast<int*>(_a[0]) = -1;
        _id -= 4;
    }
    return _id;
}

// SIGNAL 0
void ES2EvaluationResultAnalyse::dataOperate()
{
    QMetaObject::activate(this, &staticMetaObject, 0, nullptr);
}
QT_WARNING_POP
QT_END_MOC_NAMESPACE
