package com.github.segu23.xlsxconverter.exception;

import com.github.segu23.xlsxconverter.model.ExcelFieldConversionError;
import lombok.Getter;

import java.util.List;

@Getter
public class FieldsConversionException extends Exception {

    private final List<ExcelFieldConversionError> errors;

    public FieldsConversionException(String message, List<ExcelFieldConversionError> errors) {
        super(message);
        this.errors = errors;
    }
}
