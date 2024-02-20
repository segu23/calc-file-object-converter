package com.github.segu23.xlsxconverter.controller;

import com.github.segu23.xlsxconverter.CalcFileConverter;
import com.github.segu23.xlsxconverter.exception.SheetNotFoundException;
import com.github.segu23.xlsxconverter.model.SecondTestUserModel;
import com.github.segu23.xlsxconverter.model.TestUserModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
@RequestMapping("/document")
public class DocumentController {

    @PostMapping("/upload-test")
    public ResponseEntity<?> uploadDocument(
            @RequestParam("file") MultipartFile file,
            @RequestParam("startRow") int startRow,
            @RequestParam("startColumn") int startColumn,
            @RequestParam(value = "sheet", required = false) String sheetName,
            @RequestParam(value = "sheetIndex", required = false, defaultValue = "0") int sheetIndex) {
        try {
            Workbook workbook = CalcFileConverter.getWorkbook(file);
            Sheet sheet;

            if (sheetName != null && !sheetName.equalsIgnoreCase("")) {
                sheet = CalcFileConverter.getSheetByName(workbook, sheetName);
            } else {
                sheet = CalcFileConverter.getSheetByIndex(workbook, sheetIndex);
            }

            if (sheet == null) throw new SheetNotFoundException();

            int endRow = CalcFileConverter.getLastRow(sheet, startRow, startColumn);
            int endColumn = CalcFileConverter.getLastColumn(sheet, startRow, startColumn);

            List<TestUserModel> users = CalcFileConverter.extractObjectsFromTable(sheet, startRow, endRow, startColumn, endColumn, TestUserModel.class);
            users.forEach(System.out::println);

            return ResponseEntity.ok().build();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ResponseEntity.badRequest().build();
    }

    @PostMapping("/upload-second-test")
    public ResponseEntity<?> uploadDocument2(
            @RequestParam("file") MultipartFile file,
            @RequestParam("startRow") int startRow,
            @RequestParam("startColumn") int startColumn,
            @RequestParam(value = "sheet", required = false) String sheetName,
            @RequestParam(value = "sheetIndex", required = false, defaultValue = "0") int sheetIndex) {
        try {
            Workbook workbook = CalcFileConverter.getWorkbook(file);
            Sheet sheet;

            if (sheetName != null && !sheetName.equalsIgnoreCase("")) {
                sheet = CalcFileConverter.getSheetByName(workbook, sheetName);
            } else {
                sheet = CalcFileConverter.getSheetByIndex(workbook, sheetIndex);
            }

            if (sheet == null) throw new SheetNotFoundException();
            int endRow = CalcFileConverter.getLastRow(sheet, startRow, startColumn);
            int endColumn = CalcFileConverter.getLastColumn(sheet, startRow, startColumn);

            List<SecondTestUserModel> users = CalcFileConverter.extractObjectsFromTable(sheet, startRow, endRow, startColumn, endColumn, SecondTestUserModel.class);
            users.forEach(System.out::println);

            return ResponseEntity.ok().build();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return ResponseEntity.badRequest().build();
    }
}
