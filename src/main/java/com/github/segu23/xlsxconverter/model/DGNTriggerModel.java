package com.github.segu23.xlsxconverter.model;

import com.github.segu23.xlsxconverter.annotation.ExcelColumn;

public class DGNTriggerModel {

    @ExcelColumn("Name")
    private String name;
    @ExcelColumn("Cards")
    private String cards;
    @ExcelColumn("PAN_ICC")
    private String panIcc;
    @ExcelColumn("PAN_DE")
    private String panDe;
    @ExcelColumn("PAN 1")
    private Long firstPan;
    @ExcelColumn("PAN 2")
    private Long secondPan;
    @ExcelColumn("ALTERNATE PAN 1")
    private Long firstAlternatePan;
    @ExcelColumn("Track2 Data 1")
    private Long firstTrackTwoData;
    @ExcelColumn("Track2 Data 2")
    private Long secondTrackTwoData;
    @ExcelColumn("Track2 Data Alternate 1")
    private Long firstAlternateTrackTwoData;
    @ExcelColumn("AMOUNT 1")
    private String firstAmount;
    @ExcelColumn("AMOUNT 2")
    private String secondAmount;
    @ExcelColumn("PIN")
    private Integer pin;
    @ExcelColumn("Response Code")
    private Integer responseCode;

    @Override
    public String toString() {
        return "DGNMontantPanNewModel{" +
                "name='" + name + '\'' +
                ", cards='" + cards + '\'' +
                ", panIcc='" + panIcc + '\'' +
                ", panDe='" + panDe + '\'' +
                ", firstPan=" + firstPan +
                ", secondPan=" + secondPan +
                ", firstAlternatePan=" + firstAlternatePan +
                ", firstTrackTwoData=" + firstTrackTwoData +
                ", secondTrackTwoData=" + secondTrackTwoData +
                ", firstAlternateTrackTwoData=" + firstAlternateTrackTwoData +
                ", firstAmount='" + firstAmount + '\'' +
                ", secondAmount='" + secondAmount + '\'' +
                ", pin=" + pin +
                ", responseCode=" + responseCode +
                '}';
    }
}
