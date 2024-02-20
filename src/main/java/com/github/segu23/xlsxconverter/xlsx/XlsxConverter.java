package com.github.segu23.xlsxconverter.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class XlsxConverter {

    public static <T> List<T> extractObjectsFromTable(Sheet sheet, int startRow, int endRow, int startColumn, int endColumn, Class<T> type) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        List<T> objectList = new ArrayList<>();

        for (int i = startRow + 1; i <= endRow; i++) {
            String cellVal = "";

            T object = type.getDeclaredConstructor().newInstance();
            for (int j = startColumn; j <= endColumn; j++) {
                String columnName = sheet.getRow(startRow).getCell(j - startColumn).getStringCellValue();
                Cell cell = sheet.getRow(i).getCell(j);

                if (cell != null) {
                    switch (cell.getCellType()) {
                        case NUMERIC -> {
                            cellVal = String.valueOf(cell.getNumericCellValue());
                        }
                        case STRING -> {
                            cellVal = cell.getStringCellValue();
                        }
                    }

                    assignValueToField(object, columnName, cellVal);
                }


            }
            objectList.add(object);
        }

        return objectList;
    }

    private static <T> T assignValueToField(T instance, String fieldName, String cellValue) {
        try {
            Field field = instance.getClass().getDeclaredField(fieldName);

            field.setAccessible(true);

            if (field.getType() == String.class) {
                field.set(instance, cellValue);
            } else if (field.getType() == Integer.class || field.getType() == int.class) {
                field.set(instance, Double.valueOf(cellValue).intValue());
            } else if (field.getType() == Double.class || field.getType() == double.class) {
                field.set(instance, Double.valueOf(cellValue));
            } else if (field.getType() == Long.class || field.getType() == long.class) {
                field.set(instance, Double.valueOf(cellValue).longValue());
            } else if (field.getType() == LocalDate.class) {
                field.set(instance, LocalDate.of(1899, 12, 30).plusDays(Double.valueOf(cellValue).intValue()));
            }

            field.setAccessible(false);
        } catch (NoSuchFieldException | IllegalAccessException e) {
            e.printStackTrace();
        }

        return instance;
    }

    public static int getLastRow(Sheet sheet, int startRowIndex, int startColumn) {
        int endRow = startRowIndex;
        Row actualRow = sheet.getRow(endRow);

        while (actualRow != null && actualRow.getCell(startColumn).getCellType() != CellType.BLANK) {
            endRow++;
            actualRow = sheet.getRow(endRow);
        }

        return endRow-1;
    }

    public static int getLastColumn(Sheet sheet, int startRowIndex, int startColumn) {
        int endColumn = startColumn;
        Row startRow = sheet.getRow(startRowIndex);
        Cell actualColumn = startRow.getCell(endColumn);
        while (actualColumn != null && actualColumn.getCellType() != CellType.BLANK) {
            endColumn++;
            actualColumn = startRow.getCell(endColumn);
        }

        return endColumn-1;
    }
}
