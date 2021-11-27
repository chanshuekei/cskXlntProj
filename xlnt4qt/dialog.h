#ifndef DIALOG_H
#define DIALOG_H

#include <QDialog>

QT_BEGIN_NAMESPACE
namespace Ui {
    class Dialog;
}
QT_END_NAMESPACE

class Dialog : public QDialog {
    Q_OBJECT

  public:
    Dialog(QWidget *parent = nullptr);
    ~Dialog();

  private slots:
    void on_pushButton2Create_clicked();
    void on_pushButton2Write_clicked();
    void on_pushButton2Read_clicked();

  private:
    Ui::Dialog *ui;
    QString m_filename; // excel file name
};
#endif // DIALOG_H
