package com.github.segu23.xlsxconverter.model;

import com.github.segu23.xlsxconverter.annotation.ExcelColumn;

public class TestUserModel {

    @ExcelColumn("ID")
    private int id;
    @ExcelColumn("Name")
    private String name;
    @ExcelColumn("Email")
    private String email;

    @Override
    public String toString() {
        return "TestUserModel{" +
                "ID=" + id +
                ", Name='" + name + '\'' +
                ", Email='" + email + '\'' +
                '}';
    }
}
