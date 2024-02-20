package com.github.segu23.xlsxconverter.model;

import java.time.LocalDate;
import java.util.Date;

public class SecondTestUserModel {

    private int ID;
    private Long Phone;
    private LocalDate Date;

    @Override
    public String toString() {
        return "SecondTestUserModel{" +
                "ID=" + ID +
                ", Phone=" + Phone +
                ", Date=" + Date +
                '}';
    }
}
