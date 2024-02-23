package com.github.segu23.xlsxconverter.converter;

public interface DataTypeConversion<T> {

    T convert(String data);
}
