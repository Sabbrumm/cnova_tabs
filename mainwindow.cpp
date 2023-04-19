
#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QMessageBox>
#include <QTimer>
#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{



    ui->setupUi(this);

    ui->group_edit->setValidator( new QRegularExpressionValidator(QRegularExpression("\\d{4}"), this) );

    ui->tabWidget->setTabEnabled(1, false);

    QTimer *t = new QTimer(this);

    QString time1 = QTime::currentTime().toString();
    ui->label_clock->setText(time1);


    t->setInterval(1000);
    MainWindow::connect(t, &QTimer::timeout, [&]() {
        QString time1 = QTime::currentTime().toString();
        if (QTime::currentTime().hour()==0 && QTime::currentTime().minute()==0 && QTime::currentTime().second()==0)
            ui->calendarWidget->setSelectedDate(QDate::currentDate());
        ui->label_clock->setText(time1);
    } );
    t->start();

}



MainWindow::~MainWindow()
{
    delete ui;
}




void MainWindow::on_button_tostud_clicked()
{
    if (ui->fio_edit->text().isEmpty() || ui->group_edit->text().isEmpty()){
        QMessageBox::warning(this, "Внимание!", "Хотя бы одно из полей не заполнено. Необходимо заполнить все поля.");
        return;
    }

    if (ui->group_edit->text().length()!=4){
        QMessageBox::warning(this, "Внимание!", "Имя группы должно состоять из 4-х цифр!");
        return;
    }
    ui->fio_edit->setReadOnly(1);
    ui->group_edit->setReadOnly(1);
    ui->tabWidget->setTabEnabled(1, 1);
    ui->gender_male_radio->setEnabled(0);
    ui->gender_female_radio->setEnabled(0);
    ui->tabWidget->setCurrentIndex(1);
    ui->tabWidget->setTabEnabled(0, 0);

}


void MainWindow::on_getback_button_clicked()
{
    ui->fio_edit->setReadOnly(0);
    ui->group_edit->setReadOnly(0);
    ui->tabWidget->setTabEnabled(0, 1);
    ui->fio_edit->setText("");
    ui->group_edit->setText("");
    ui->gender_male_radio->setChecked(1);
    ui->gender_male_radio->setEnabled(1);
    ui->gender_female_radio->setEnabled(1);
    ui->tabWidget->setCurrentIndex(0);
    ui->tabWidget->setTabEnabled(1, 0);
    while (ui->tableWidget->rowCount()){
        ui->tableWidget->removeRow(ui->tableWidget->rowCount()-1);
    }

}


void MainWindow::on_addrow_button_clicked()
{
    if (ui->tableWidget->rowCount()<10)
        ui->tableWidget->insertRow(ui->tableWidget->currentRow()+1);
    else
        QMessageBox::critical(this, "О БОЖЕ!", "Больше 10 экзаменов? Вы действительно так замучали этого бедного студента?");
}


void MainWindow::on_delrow_button_clicked()
{
    ui->tableWidget->removeRow((ui->tableWidget->currentRow()>=0)?(ui->tableWidget->currentRow()):(ui->tableWidget->rowCount()-1));
}



void MainWindow::on_sorting_button_clicked()
{
    ui->tableWidget->sortByColumn(ui->sortBox->currentIndex(), Qt::DescendingOrder);
}


void MainWindow::on_export_button_clicked()
{
    QString name = ui->fio_edit->text().replace(" ", "_") + "_"+ ui->group_edit->text();
    QString filter = "Лист Microsoft Excel (*.xlsx)";
    QString filepath = QFileDialog::getSaveFileName(this, "Сохранение файла", "C:\\Users\\" + name, filter, &filter);
    if (!filepath.endsWith(".xlsx")) {
        filepath += ".xlsx";
    }

    int rowCount = ui->tableWidget->rowCount();

    QAxObject *excel = new QAxObject(this);
    excel->setControl("Excel.Application");
    excel->dynamicCall("SetVisible (bool Visible)", "false");
    excel->setProperty("DisplayAlerts", false);

    QAxObject *workbooks = excel->querySubObject("WorkBooks");
    workbooks->dynamicCall("Add");
    QAxObject *workbook = excel->querySubObject("ActiveWorkBook");
    QAxObject *worksheets = workbook->querySubObject("Sheets");

    int sheetsCount = worksheets->property("Count").toInt();

    QAxObject* sheet1 = worksheets->querySubObject("Item(int)", sheetsCount);
    worksheets->dynamicCall("Add(QVariant)", sheet1->asVariant());

    QAxObject* sheet2 = worksheets->querySubObject("Item(int)", sheetsCount);
    sheet1->dynamicCall("Move(QVariant)", sheet2->asVariant());

    sheet1->dynamicCall("SetName(const QVariant&)", QVariant("Оценки"));
    sheet2->dynamicCall("SetName(const QVariant&)", QVariant("Студент"));

    //Заполнение оценок

    QAxObject *cellA, *cellB;
    QString columnA = "A1";
    QString columnB = "B1";
    cellA = sheet1->querySubObject("Range(QVariant, QVariant)", columnA);
    cellB = sheet1->querySubObject("Range(QVariant, QVariant)", columnB);

    cellA->dynamicCall("SetValue(const QVariant&)", QVariant(ui->tableWidget->horizontalHeaderItem(0)->text()));
    cellB->dynamicCall("SetValue(const QVariant&)", QVariant(ui->tableWidget->horizontalHeaderItem(1)->text()));


    for (int i = 0; i < rowCount; i++) {
        columnA = "A" + QString::number(i + 2);
        columnB = "B" + QString::number(i + 2);

        cellA = sheet1->querySubObject("Range(QVariant, QVariant)", columnA);
        cellB = sheet1->querySubObject("Range(QVariant, QVariant)", columnB);

        qDebug() <<ui->tableWidget->item(i, 0)->text();
        qDebug() <<ui->tableWidget->item(i, 1)->text();

        cellA->dynamicCall("SetValue(const QVariant&)", QVariant(ui->tableWidget->item(i, 0)->text()));
        cellB->dynamicCall("SetValue(const QVariant&)", QVariant(ui->tableWidget->item(i, 1)->text()));
    }

    //Заполнение инфы о студенте
    QAxObject *cell;

    cell = sheet2->querySubObject("Range(QVariant, QVariant)", QString("A1"));
    cell->dynamicCall("SetValue(const QVariant&)", QVariant("ФИО:"));

    cell = sheet2->querySubObject("Range(QVariant, QVariant)", QString("B1"));
    cell->dynamicCall("SetValue(const QVariant&)", QVariant(ui->fio_edit->text()));


    cell = sheet2->querySubObject("Range(QVariant, QVariant)", QString("A2"));
    cell->dynamicCall("SetValue(const QVariant&)", QVariant("Группа:"));

    cell = sheet2->querySubObject("Range(QVariant, QVariant)", QString("B2"));
    cell->dynamicCall("SetValue(const QVariant&)", QVariant(ui->group_edit->text()));


    cell = sheet2->querySubObject("Range(QVariant, QVariant)", QString("A3"));
    cell->dynamicCall("SetValue(const QVariant&)", QVariant("Пол:"));

    cell = sheet2->querySubObject("Range(QVariant, QVariant)", QString("B3"));
    cell->dynamicCall("SetValue(const QVariant&)",
        QVariant(
                    (ui->gender_male_radio->isChecked()) ? ("Мужской"):("Женский")
                )
    );

    workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(filepath));
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
}

