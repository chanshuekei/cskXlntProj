#include "dialog.h"
#include "ui_dialog.h"

#include <QDateTime>
#include <QDebug>
#include "xlnt/xlnt.hpp"

#ifndef X
#define X(x) QString::fromLocal8Bit(x)
#endif

Dialog::Dialog(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::Dialog) {
    ui->setupUi(this);
    setWindowTitle("TestXlnt");
}

Dialog::~Dialog() {
    delete ui;
}


void Dialog::on_pushButton2Create_clicked() {
    m_filename = ui->lineEdit2Create->text().trimmed();

    if (m_filename.isEmpty()) {
        ui->plainTextEdit->appendPlainText("File name can't to be empty!");
        return;
    }

    if (!m_filename.endsWith(".xlsx")) {
        ui->plainTextEdit->appendPlainText("File is not excel file!");
        m_filename.clear();
        return;
    }

    xlnt::workbook().save(m_filename.toStdString());
    ui->plainTextEdit->appendPlainText(X("Create excel file Succ:%1").arg(m_filename));
}


void Dialog::on_pushButton2Write_clicked() {
    if (m_filename.isEmpty()) {
        ui->plainTextEdit->appendPlainText(X("Can not file the excel file"));
        return;
    }

    try {
        xlnt::workbook book;
        book.load(m_filename.toStdString());
        xlnt::worksheet sheet = book.active_sheet();
        xlnt::cell a1 = sheet.cell("A1");
        xlnt::cell b1 = sheet.cell("B1");
        xlnt::cell c1 = sheet.cell("C1");
        xlnt::cell d1 = sheet.cell("D1");
        // add record
        QString a1Data = X("Test");
        QString b1Data = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
        QString c1Data = X("[Plate]:%1\n[Color]:%2\n[Hight]:%3\n[Length]:%4\n[Wide]:%5\n")
                         .arg(X("T9413"))
                         .arg(X("Red"))
                         .arg(200)
                         .arg(450)
                         .arg(220);
        QString d1Data = X("[Plate]:%1\n[Color]:%2\n[Hight]:%3\n[Length]:%4\n[Wide]:%5\n")
                         .arg(X("N9527"))
                         .arg(X("BLUE"))
                         .arg(220)
                         .arg(460)
                         .arg(240);
        ui->plainTextEdit->appendPlainText(a1Data);
        ui->plainTextEdit->appendPlainText(b1Data);
        ui->plainTextEdit->appendPlainText(c1Data);
        ui->plainTextEdit->appendPlainText(d1Data);
        // style
        xlnt::alignment alig = xlnt::alignment()
                               .wrap(true)
                               .horizontal(xlnt::horizontal_alignment::left)
                               .vertical(xlnt::vertical_alignment::center);
        a1.alignment(alig);
        b1.alignment(alig);
        c1.alignment(alig);
        d1.alignment(alig);
        // font size:12
        xlnt::font font = xlnt::font().size(12);
        a1.font(font);
        b1.font(font);
        c1.font(font);
        d1.font(font);
        // save data
        a1.value(a1Data.toStdString());
        b1.value(b1Data.toStdString());
        c1.value(c1Data.toStdString());
        d1.value(d1Data.toStdString());
        // border
        xlnt::border bd;
        xlnt::border::border_property bp;
        bp.color(xlnt::color::black());
        bp.style(xlnt::border_style::medium);

        for (auto side :  {
                    xlnt::border_side::bottom,
                    xlnt::border_side::end,
                    xlnt::border_side::horizontal,
                    xlnt::border_side::start,
                    xlnt::border_side::top,
                    xlnt::border_side::vertical
                })
            bd.side(side, bp);
        xlnt::range range = sheet.range({a1.reference(), d1.reference()});
        range.border(bd);
        book.save(m_filename.toStdString());
        ui->plainTextEdit->appendPlainText(X("Saved"));
    } catch (std::exception e) {
        ui->plainTextEdit->appendPlainText(X("Error:") + e.what());
    }

    ui->plainTextEdit->appendPlainText(X("##############################"));
}


void Dialog::on_pushButton2Read_clicked() {
    if (m_filename.isEmpty()) {
        ui->plainTextEdit->appendPlainText(X("Can not file the excel file"));
        return;
    }

    xlnt::workbook book;
    book.load(m_filename.toStdString());
    xlnt::worksheet sheet = book.active_sheet();

    for (auto row : sheet.rows(false)) {
        for (auto cell : row) {
            QString val = QString::fromStdString(cell.to_string());
            ui->plainTextEdit->appendPlainText(val);
        }
    }

    ui->plainTextEdit->appendPlainText(X("##############################"));
}

