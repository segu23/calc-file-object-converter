package com.github.segu23.xlsxconverter.converter;

import com.github.segu23.xlsxconverter.annotation.ExcelColumn;
import com.github.segu23.xlsxconverter.exception.DataTypeConversionNotFoundException;
import com.github.segu23.xlsxconverter.exception.ExcelFieldNotFoundException;
import com.github.segu23.xlsxconverter.exception.SheetNotFoundException;
import jakarta.annotation.Nullable;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    public static <T> List<T> extractObjectsFromTable(Sheet sheet, int startRow, int endRow, int startColumn, int endColumn, Class<T> type) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, ExcelFieldNotFoundException, DataTypeConversionNotFoundException {
        String[] columnNames = getTableColumnNames(sheet, startRow, startColumn, endColumn);

        return extractObjectsFromTable(sheet, startRow, endRow, startColumn, endColumn, type, columnNames);
    }

    private static String[] getTableColumnNames(Sheet sheet, int startRow, int startColumn, int endColumn) {
        String[] columnNames = new String[endColumn - startColumn + 1];

        for (int j = startColumn; j <= endColumn; j++) {
            String columnName = sheet.getRow(startRow).getCell(j).getStringCellValue();
            columnNames[j - startColumn] = columnName;
        }

        return columnNames;
    }

    public static <T> List<T> extractObjectsFromTable(Sheet sheet, int startRow, int endRow, int startColumn, int endColumn, Class<T> type, String[] columns) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, ExcelFieldNotFoundException, DataTypeConversionNotFoundException {
        Map<String, Field> classFields = getClassExcelColumnFieldsMap(type);
        List<T> objectList = new ArrayList<>();

        for (int i = startRow + 1; i <= endRow; i++) {
            T object = type.getDeclaredConstructor().newInstance();

            for (int j = startColumn; j <= endColumn; j++) {
                String columnName = columns[j - startColumn];
                Cell cell = sheet.getRow(i).getCell(j);
                String cellVal = "";

                if (cell != null) {
                    switch (cell.getCellType()) {
                        case BLANK -> {
                            continue;
                        }
                        case NUMERIC -> {
                            cellVal = String.valueOf(cell.getNumericCellValue());
                        }
                        case STRING -> {
                            cellVal = cell.getStringCellValue();
                        }
                        case BOOLEAN -> {
                            cellVal = String.valueOf(cell.getBooleanCellValue());
                        }
                        case FORMULA -> {
                            switch (cell.getCachedFormulaResultType()) {
                                case STRING -> {
                                    cellVal = cell.getStringCellValue();
                                }
                                case NUMERIC -> {
                                    cellVal = String.valueOf(cell.getNumericCellValue());
                                }
                                case BOOLEAN -> {
                                    cellVal = String.valueOf(cell.getBooleanCellValue());
                                }
                            }
                        }
                    }

                    assignValueToField(object, columnName, cellVal, classFields);
                }
            }

            objectList.add(object);
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

    public static Sheet getSheetByName(Workbook workbook, @Nullable String sheetName) throws SheetNotFoundException {
        Sheet sheet = workbook.getSheet(sheetName);

        if (sheet == null) throw new SheetNotFoundException();

        return sheet;
    }

    public static Workbook getWorkbook(MultipartFile file) throws IOException {
        Workbook workbook;
        workbook = WorkbookFactory.create(file.getInputStream());

//        if (file.getOriginalFilename().endsWith(".xlsx")) {
//            workbook = new XSSFWorkbook(file.getInputStream());
//        } else if (file.getOriginalFilename().endsWith(".xls")) {
//            workbook = new HSSFWorkbook(file.getInputStream());
//        } else if (file.getOriginalFilename().endsWith(".xlsm")) {
//            workbook = WorkbookFactory.create(file.getInputStream());
//        } else {
//            throw new InvalidFileTypeException();
//        }

        return workbook;
    }
}