package com.github.segu23.xlsxconverter.model;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class ExcelObjectWrapper<T> {

    private int row;

    private T object;
}
