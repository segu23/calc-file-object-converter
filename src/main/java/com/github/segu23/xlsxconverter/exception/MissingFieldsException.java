package com.github.segu23.xlsxconverter.exception;

import lombok.Getter;

import java.util.List;

@Getter
public class MissingFieldsException extends Exception {

    private final List<String> missingFields;

    public MissingFieldsException(List<String> missingFields) {
        super("Missing fields: " + String.join(", ", missingFields));
        this.missingFields = missingFields;
    }
}
