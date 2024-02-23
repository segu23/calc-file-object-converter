package com.github.segu23.xlsxconverter.model;

import com.github.segu23.xlsxconverter.annotation.ExcelColumn;

public class DGNPassCriteriaIndexModel {

    @ExcelColumn("Index")
    private Integer index;
    @ExcelColumn("Criteria")
    private CriteriaModel criteriaModel;

    @Override
    public String toString() {
        return "DGNPassCriteriaIndexModel{" +
                "index=" + index +
                ", criteriaModel=" + criteriaModel +
                '}';
    }
}
