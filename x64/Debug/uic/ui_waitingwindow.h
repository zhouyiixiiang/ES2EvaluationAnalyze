/********************************************************************************
** Form generated from reading UI file 'waitingwindow.ui'
**
** Created by: Qt User Interface Compiler version 5.12.1
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_WAITINGWINDOW_H
#define UI_WAITINGWINDOW_H

#include <QtCore/QVariant>
#include <QtWidgets/QApplication>
#include <QtWidgets/QFrame>
#include <QtWidgets/QHBoxLayout>
#include <QtWidgets/QLabel>
#include <QtWidgets/QVBoxLayout>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_WaitingWindow
{
public:
    QVBoxLayout *verticalLayout;
    QFrame *frame;
    QHBoxLayout *horizontalLayout;
    QLabel *label;

    void setupUi(QWidget *WaitingWindow)
    {
        if (WaitingWindow->objectName().isEmpty())
            WaitingWindow->setObjectName(QString::fromUtf8("WaitingWindow"));
        WaitingWindow->resize(400, 300);
        WaitingWindow->setAutoFillBackground(true);
        verticalLayout = new QVBoxLayout(WaitingWindow);
        verticalLayout->setSpacing(6);
        verticalLayout->setContentsMargins(11, 11, 11, 11);
        verticalLayout->setObjectName(QString::fromUtf8("verticalLayout"));
        frame = new QFrame(WaitingWindow);
        frame->setObjectName(QString::fromUtf8("frame"));
        frame->setAutoFillBackground(true);
        frame->setFrameShape(QFrame::StyledPanel);
        frame->setFrameShadow(QFrame::Raised);
        horizontalLayout = new QHBoxLayout(frame);
        horizontalLayout->setSpacing(6);
        horizontalLayout->setContentsMargins(11, 11, 11, 11);
        horizontalLayout->setObjectName(QString::fromUtf8("horizontalLayout"));
        label = new QLabel(frame);
        label->setObjectName(QString::fromUtf8("label"));
        QFont font;
        font.setPointSize(14);
        label->setFont(font);
        label->setAutoFillBackground(true);
        label->setFrameShape(QFrame::Panel);

        horizontalLayout->addWidget(label);


        verticalLayout->addWidget(frame);


        retranslateUi(WaitingWindow);

        QMetaObject::connectSlotsByName(WaitingWindow);
    } // setupUi

    void retranslateUi(QWidget *WaitingWindow)
    {
        WaitingWindow->setWindowTitle(QApplication::translate("WaitingWindow", "\347\255\211\345\276\205\344\270\255", nullptr));
        label->setText(QApplication::translate("WaitingWindow", "\346\255\243\345\234\250\350\257\273\345\206\231\346\225\260\346\215\256\357\274\214\345\244\204\347\220\206\347\254\2540/0\344\270\252\346\225\260\346\215\256", nullptr));
    } // retranslateUi

};

namespace Ui {
    class WaitingWindow: public Ui_WaitingWindow {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_WAITINGWINDOW_H
