package com.github.segu23.xlsxconverter.model;

import com.github.segu23.xlsxconverter.annotation.ExcelColumn;

import java.time.LocalDate;

public class SecondTestUserModel {

    @ExcelColumn("ID")
    private int id;
    @ExcelColumn("Phone")
    private Long phone;
    @ExcelColumn("Date")
    private LocalDate date;

    @Override
    public String toString() {
        return "SecondTestUserModel{" +
                "ID=" + id +
                ", Phone=" + phone +
                ", Date=" + date +
                '}';
    }
}
