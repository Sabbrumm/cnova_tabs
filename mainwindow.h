
#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>



QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow

{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:

    void on_button_tostud_clicked();

    void on_getback_button_clicked();

    void on_addrow_button_clicked();

    void on_delrow_button_clicked();

    void on_sorting_button_clicked();

    void on_export_button_clicked();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
