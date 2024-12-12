package com.github.segu23.xlsxconverter.exception;

import com.github.segu23.xlsxconverter.model.ExcelFieldConversionError;
import com.github.segu23.xlsxconverter.model.ExcelObjectWrapper;
import lombok.Getter;

import java.util.List;

@Getter
public class FieldsConversionException extends Exception {

    private final List<ExcelFieldConversionError> errors;
    private final List<ExcelObjectWrapper<Object>> correctObjects;

    public FieldsConversionException(String message, List<ExcelFieldConversionError> errors, List<ExcelObjectWrapper<Object>> correctObjects) {
        super(message);
        this.errors = errors;
        this.correctObjects = correctObjects;
    }
}
