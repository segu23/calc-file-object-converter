package com.github.segu23.xlsxconverter.model;

public class CriteriaModel {

    private String name;
    private String value;

    public CriteriaModel(String name, String value) {
        this.name = name;
        this.value = value;
    }

    @Override
    public String toString() {
        return "CriteriaModel{" +
                "name='" + name + '\'' +
                ", value='" + value + '\'' +
                '}';
    }
}
