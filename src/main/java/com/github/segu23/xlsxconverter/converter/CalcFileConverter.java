package com.github.segu23.xlsxconverter.converter;

import com.github.segu23.xlsxconverter.annotation.ExcelColumn;
import com.github.segu23.xlsxconverter.exception.*;
import com.github.segu23.xlsxconverter.model.ExcelFieldConversionError;
import com.github.segu23.xlsxconverter.model.ExcelObjectWrapper;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

public class CalcFileConverter {

    public static final Map<Class<?>, DataTypeConversion<?>> DATA_TYPE_CONVERSIONS = new HashMap<>();

    static {
        addDataTypeConversion(String.class, data -> data);
        addDataTypeConversion(Integer.class, data -> Double.valueOf(data).intValue());
        addDataTypeConversion(int.class, data -> Double.valueOf(data).intValue());
        addDataTypeConversion(Double.class, Double::valueOf);
        addDataTypeConversion(double.class, Double::valueOf);
        addDataTypeConversion(Long.class, data -> Double.valueOf(data).longValue());
        addDataTypeConversion(long.class, data -> Double.valueOf(data).longValue());
        addDataTypeConversion(LocalDate.class, data -> LocalDate.of(1899, 12, 30).plusDays(Double.valueOf(data).intValue()));
        addDataTypeConversion(Boolean.class, Boolean::valueOf);
        addDataTypeConversion(boolean.class, Boolean::valueOf);
    }

    public static void addDataTypeConversion(Class<?> type, DataTypeConversion<?> dataTypeConversion) {
        DATA_TYPE_CONVERSIONS.put(type, dataTypeConversion);
    }

    public static <T> List<ExcelObjectWrapper<T>> extractObjectsFromTable(Sheet sheet, int startRow, int endRow, int startColumn, int endColumn, Class<T> type) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, ExcelFieldNotFoundException, DataTypeConversionNotFoundException, FieldsConversionException, MissingFieldsException {
        String[] columnNames = getTableColumnNames(sheet, startRow, startColumn, endColumn);

        return extractObjectsFromTable(sheet, startRow + 1, endRow, startColumn, endColumn, type, columnNames);
    }

    private static String[] getTableColumnNames(Sheet sheet, int startRow, int startColumn, int endColumn) {
        String[] columnNames = new String[endColumn - startColumn + 1];

        for (int j = startColumn; j <= endColumn; j++) {
            if (sheet.getRow(startRow) == null || sheet.getRow(startRow).getCell(j) == null) {
                columnNames[j - startColumn] = null;
                continue;
            }
            String columnName = sheet.getRow(startRow).getCell(j).getStringCellValue();
            columnNames[j - startColumn] = columnName;
        }

        return columnNames;
    }

    public static <T> List<ExcelObjectWrapper<T>> extractObjectsFromTable(Sheet sheet, int startRow, int endRow, int startColumn, int endColumn, Class<T> type, String[] columns) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, FieldsConversionException, MissingFieldsException {
        Map<String, Field> classFields = getClassExcelColumnFieldsMap(type);
        List<String> missingFields = new ArrayList<>();
        List<String> columnsList = Arrays.asList(columns);
        classFields.forEach((key, value) -> {
            if (!columnsList.contains(key)) {
                missingFields.add(key);
            }
        });
        if (!missingFields.isEmpty()) {
            throw new MissingFieldsException(missingFields);
        }

        List<ExcelObjectWrapper<T>> objectList = new ArrayList<>();
        List<ExcelFieldConversionError> errors = new ArrayList<>();

        for (int i = startRow; i <= endRow; i++) {
            T object = type.getDeclaredConstructor().newInstance();
            boolean errorsFound = false;

            for (int j = startColumn; j <= endColumn; j++) {
                String columnName = columns[j - startColumn];
                Cell cell = sheet.getRow(i).getCell(j);
                String cellVal = "";

                if (cell != null) {
                    switch (cell.getCellType()) {
                        case BLANK: {
                            continue;
                        }
                        case NUMERIC: {
                            cellVal = String.valueOf(cell.getNumericCellValue());
                            break;
                        }
                        case STRING: {
                            cellVal = cell.getStringCellValue();
                            break;
                        }
                        case BOOLEAN: {
                            cellVal = String.valueOf(cell.getBooleanCellValue());
                            break;
                        }
                        case FORMULA: {
                            switch (cell.getCachedFormulaResultType()) {
                                case STRING: {
                                    cellVal = cell.getStringCellValue();
                                    break;
                                }
                                case NUMERIC: {
                                    cellVal = String.valueOf(cell.getNumericCellValue());
                                    break;
                                }
                                case BOOLEAN: {
                                    cellVal = String.valueOf(cell.getBooleanCellValue());
                                    break;
                                }
                            }
                            break;
                        }
                    }

                    try {
                        assignValueToField(object, columnName, cellVal, classFields);
                    } catch (Exception e) {
                        errors.add(new ExcelFieldConversionError(i, columnName, cellVal));
                        errorsFound = true;
                    }
                }
            }

            if (errorsFound) {
                continue;
            }
            objectList.add(new ExcelObjectWrapper<>(i, object));
        }

        if (!errors.isEmpty()) {
            throw new FieldsConversionException("Error converting fields", errors, new ArrayList<>(objectList.stream().map(obj -> (ExcelObjectWrapper<Object>) obj).collect(Collectors.toList())));
        }

        return objectList;
    }

    private static <T> Map<String, Field> getClassExcelColumnFieldsMap(Class<T> type) {
        Map<String, Field> classFields = new HashMap<>();

        for (Field field : type.getDeclaredFields()) {
            field.setAccessible(true);

            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                String excelFieldName = excelColumn.value();
                classFields.put(excelFieldName, field);
            }
        }

        return classFields;
    }

    private static <T> void assignValueToField(T instance, String fieldName, String cellValue, Map<String, Field> classFields) throws DataTypeConversionNotFoundException, ExcelFieldNotFoundException, IllegalAccessException {
        Field field = classFields.get(fieldName);

        if (field == null) throw new ExcelFieldNotFoundException(fieldName);

        field.setAccessible(true);

        DataTypeConversion<?> dataTypeConversion = DATA_TYPE_CONVERSIONS.get(field.getType());

        if (dataTypeConversion == null) throw new DataTypeConversionNotFoundException(field.getType().getName());

        field.set(instance, dataTypeConversion.convert(cellValue));
    }

    public static int getLastRow(Sheet sheet, int startRowIndex, int startColumn) {
        int endRow = startRowIndex;
        Row actualRow = sheet.getRow(endRow);

        while (actualRow != null && actualRow.getCell(startColumn) != null && actualRow.getCell(startColumn).getCellType() != CellType.BLANK) {
            endRow++;
            actualRow = sheet.getRow(endRow);
        }

        return endRow - 1;
    }

    public static int getLastColumn(Sheet sheet, int startRowIndex, int startColumn) {
        int endColumn = startColumn;
        Row startRow = sheet.getRow(startRowIndex);
        Cell actualColumn = startRow.getCell(endColumn);
        while (actualColumn != null && actualColumn.getCellType() != CellType.BLANK) {
            endColumn++;
            actualColumn = startRow.getCell(endColumn);
        }

        return endColumn - 1;
    }

    public static Sheet getSheetByIndex(Workbook workbook, int sheetIndex) throws SheetNotFoundException {
        if (sheetIndex >= workbook.getNumberOfSheets()) throw new SheetNotFoundException();

        return workbook.getSheetAt(sheetIndex);
    }

    public static Sheet getSheetByName(Workbook workbook, String sheetName) throws SheetNotFoundException {
        Sheet sheet = workbook.getSheet(sheetName);

        if (sheet == null) throw new SheetNotFoundException();

        return sheet;
    }

    public static Workbook getWorkbook(InputStream file) throws IOException {
        Workbook workbook;
        workbook = WorkbookFactory.create(file);

        return workbook;
    }
}