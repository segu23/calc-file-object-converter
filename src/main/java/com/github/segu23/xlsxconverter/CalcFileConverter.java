package com.github.segu23.xlsxconverter;

import com.github.segu23.xlsxconverter.exception.InvalidFileTypeException;
import com.github.segu23.xlsxconverter.exception.SheetNotFoundException;
import jakarta.annotation.Nullable;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class CalcFileConverter {

    public static <T> List<T> extractObjectsFromTable(Sheet sheet, int startRow, int endRow, int startColumn, int endColumn, Class<T> type) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        List<T> objectList = new ArrayList<>();

        for (int i = startRow + 1; i <= endRow; i++) {
            String cellVal = "";

            T object = type.getDeclaredConstructor().newInstance();
            for (int j = startColumn; j <= endColumn; j++) {
                String columnName = sheet.getRow(startRow).getCell(j).getStringCellValue();
                Cell cell = sheet.getRow(i).getCell(j);

                if (cell != null) {
                    switch (cell.getCellType()) {
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

            if (field.getType().equals(String.class)) {
                field.set(instance, cellValue);
            } else if (field.getType().equals(Integer.class) || field.getType().equals(int.class)) {
                field.set(instance, Double.valueOf(cellValue).intValue());
            } else if (field.getType().equals(Double.class) || field.getType().equals(double.class)) {
                field.set(instance, Double.valueOf(cellValue));
            } else if (field.getType().equals(Long.class) || field.getType().equals(long.class)) {
                field.set(instance, Double.valueOf(cellValue).longValue());
            } else if (field.getType().equals(LocalDate.class)) {
                field.set(instance, LocalDate.of(1899, 12, 30).plusDays(Double.valueOf(cellValue).intValue()));
            } else if (field.getType().equals(Boolean.class) || field.getType().equals(boolean.class)) {
                field.set(instance, Boolean.valueOf(cellValue));
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
        if(sheetIndex >= workbook.getNumberOfSheets()) throw new SheetNotFoundException();

        Sheet sheet = workbook.getSheetAt(sheetIndex);

        return sheet;
    }

    public static Sheet getSheetByName(Workbook workbook, @Nullable String sheetName) throws IOException, InvalidFileTypeException, SheetNotFoundException {
        Sheet sheet = workbook.getSheet(sheetName);

        if(sheet == null) throw new SheetNotFoundException();

        return sheet;
    }

    public static Workbook getWorkbook(MultipartFile file) throws IOException, InvalidFileTypeException {
        Workbook workbook;

        if(file.getOriginalFilename().endsWith(".xlsx")){
            workbook = new XSSFWorkbook(file.getInputStream());
        }else if(file.getOriginalFilename().endsWith(".xls")){
            workbook = new HSSFWorkbook(file.getInputStream());
        }else{
            throw new InvalidFileTypeException();
        }

        return workbook;
    }
}
