package com.github.segu23.xlsxconverter.model;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class ExcelFieldConversionError {

    private int row;

    private String field;

    private String cellValue;
}
