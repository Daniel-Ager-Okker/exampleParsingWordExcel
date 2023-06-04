/********************************************************************************
** Form generated from reading UI file 'main_window.ui'
**
** Created by: Qt User Interface Compiler version 5.15.2
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_MAIN_WINDOW_H
#define UI_MAIN_WINDOW_H

#include <QtCore/QVariant>
#include <QtWidgets/QApplication>
#include <QtWidgets/QCheckBox>
#include <QtWidgets/QHBoxLayout>
#include <QtWidgets/QLabel>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QPushButton>
#include <QtWidgets/QSpacerItem>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QVBoxLayout>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_ConverterWindow
{
public:
    QWidget *centralwidget;
    QHBoxLayout *horizontalLayout_4;
    QVBoxLayout *verticalLayout;
    QHBoxLayout *horizontalLayout;
    QPushButton *pbLoadOldWord;
    QCheckBox *chbxExportOld2Excel;
    QLabel *lblOldWord;
    QHBoxLayout *horizontalLayout_2;
    QPushButton *pbLoadNewWord;
    QCheckBox *chbxExportNew2Excel;
    QLabel *lblNewWord;
    QPushButton *pbCompareAndExport;
    QPushButton *pbClean;
    QSpacerItem *horizontalSpacer;
    QMenuBar *menubar;
    QStatusBar *statusbar;

    void setupUi(QMainWindow *ConverterWindow)
    {
        if (ConverterWindow->objectName().isEmpty())
            ConverterWindow->setObjectName(QString::fromUtf8("ConverterWindow"));
        ConverterWindow->resize(820, 252);
        centralwidget = new QWidget(ConverterWindow);
        centralwidget->setObjectName(QString::fromUtf8("centralwidget"));
        horizontalLayout_4 = new QHBoxLayout(centralwidget);
        horizontalLayout_4->setObjectName(QString::fromUtf8("horizontalLayout_4"));
        verticalLayout = new QVBoxLayout();
        verticalLayout->setObjectName(QString::fromUtf8("verticalLayout"));
        horizontalLayout = new QHBoxLayout();
        horizontalLayout->setObjectName(QString::fromUtf8("horizontalLayout"));
        pbLoadOldWord = new QPushButton(centralwidget);
        pbLoadOldWord->setObjectName(QString::fromUtf8("pbLoadOldWord"));
        pbLoadOldWord->setMinimumSize(QSize(150, 23));
        pbLoadOldWord->setMaximumSize(QSize(150, 23));

        horizontalLayout->addWidget(pbLoadOldWord);

        chbxExportOld2Excel = new QCheckBox(centralwidget);
        chbxExportOld2Excel->setObjectName(QString::fromUtf8("chbxExportOld2Excel"));
        chbxExportOld2Excel->setMaximumSize(QSize(100, 16777215));

        horizontalLayout->addWidget(chbxExportOld2Excel);

        lblOldWord = new QLabel(centralwidget);
        lblOldWord->setObjectName(QString::fromUtf8("lblOldWord"));

        horizontalLayout->addWidget(lblOldWord);


        verticalLayout->addLayout(horizontalLayout);

        horizontalLayout_2 = new QHBoxLayout();
        horizontalLayout_2->setObjectName(QString::fromUtf8("horizontalLayout_2"));
        pbLoadNewWord = new QPushButton(centralwidget);
        pbLoadNewWord->setObjectName(QString::fromUtf8("pbLoadNewWord"));
        pbLoadNewWord->setMinimumSize(QSize(150, 23));
        pbLoadNewWord->setMaximumSize(QSize(150, 23));

        horizontalLayout_2->addWidget(pbLoadNewWord);

        chbxExportNew2Excel = new QCheckBox(centralwidget);
        chbxExportNew2Excel->setObjectName(QString::fromUtf8("chbxExportNew2Excel"));
        chbxExportNew2Excel->setMaximumSize(QSize(100, 16777215));

        horizontalLayout_2->addWidget(chbxExportNew2Excel);

        lblNewWord = new QLabel(centralwidget);
        lblNewWord->setObjectName(QString::fromUtf8("lblNewWord"));

        horizontalLayout_2->addWidget(lblNewWord);


        verticalLayout->addLayout(horizontalLayout_2);

        pbCompareAndExport = new QPushButton(centralwidget);
        pbCompareAndExport->setObjectName(QString::fromUtf8("pbCompareAndExport"));
        pbCompareAndExport->setMinimumSize(QSize(150, 23));
        pbCompareAndExport->setMaximumSize(QSize(150, 23));

        verticalLayout->addWidget(pbCompareAndExport);

        pbClean = new QPushButton(centralwidget);
        pbClean->setObjectName(QString::fromUtf8("pbClean"));
        pbClean->setMinimumSize(QSize(150, 23));
        pbClean->setMaximumSize(QSize(150, 23));

        verticalLayout->addWidget(pbClean);


        horizontalLayout_4->addLayout(verticalLayout);

        horizontalSpacer = new QSpacerItem(40, 20, QSizePolicy::Expanding, QSizePolicy::Minimum);

        horizontalLayout_4->addItem(horizontalSpacer);

        ConverterWindow->setCentralWidget(centralwidget);
        menubar = new QMenuBar(ConverterWindow);
        menubar->setObjectName(QString::fromUtf8("menubar"));
        menubar->setGeometry(QRect(0, 0, 820, 21));
        ConverterWindow->setMenuBar(menubar);
        statusbar = new QStatusBar(ConverterWindow);
        statusbar->setObjectName(QString::fromUtf8("statusbar"));
        ConverterWindow->setStatusBar(statusbar);

        retranslateUi(ConverterWindow);

        QMetaObject::connectSlotsByName(ConverterWindow);
    } // setupUi

    void retranslateUi(QMainWindow *ConverterWindow)
    {
        ConverterWindow->setWindowTitle(QCoreApplication::translate("ConverterWindow", "VelesStroyApp", nullptr));
        pbLoadOldWord->setText(QCoreApplication::translate("ConverterWindow", "LoadOldWord", nullptr));
        chbxExportOld2Excel->setText(QCoreApplication::translate("ConverterWindow", "extract To Excel", nullptr));
        lblOldWord->setText(QCoreApplication::translate("ConverterWindow", ":", nullptr));
        pbLoadNewWord->setText(QCoreApplication::translate("ConverterWindow", "LoadNewWord", nullptr));
        chbxExportNew2Excel->setText(QCoreApplication::translate("ConverterWindow", "extract To Excel", nullptr));
        lblNewWord->setText(QCoreApplication::translate("ConverterWindow", ":", nullptr));
        pbCompareAndExport->setText(QCoreApplication::translate("ConverterWindow", "CompareAndExport", nullptr));
        pbClean->setText(QCoreApplication::translate("ConverterWindow", "Clean", nullptr));
    } // retranslateUi

};

namespace Ui {
    class ConverterWindow: public Ui_ConverterWindow {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_MAIN_WINDOW_H
