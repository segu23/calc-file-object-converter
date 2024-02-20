package com.github.segu23.xlsxconverter.controller;

import com.github.segu23.xlsxconverter.exception.InvalidFileTypeException;
import com.github.segu23.xlsxconverter.exception.SheetNotFoundException;
import com.github.segu23.xlsxconverter.model.SecondTestUserModel;
import com.github.segu23.xlsxconverter.model.TestUserModel;
import com.github.segu23.xlsxconverter.xlsx.XlsxConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

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

        try{
            Sheet sheet = getSheet(file, sheetName, sheetIndex);

            int endRow = XlsxConverter.getLastRow(sheet, startRow, startColumn);
            System.out.println("EndRow: " + endRow);
            int endColumn = XlsxConverter.getLastColumn(sheet, startRow, startColumn);
            System.out.println("EndColumn: " + endColumn);

            List<TestUserModel> users = XlsxConverter.extractObjectsFromTable(sheet, startRow, endRow, startColumn, endColumn, TestUserModel.class);
            users.forEach(System.out::println);


            return ResponseEntity.ok().build();
        }catch (Exception e){
            e.printStackTrace();
        }
        return ResponseEntity.badRequest().build();
    }

    @PostMapping("/upload-second-test")
    public ResponseEntity<?> uploadDocument2(
            @RequestParam("file") MultipartFile file,
            @RequestParam("startRow") int startRow,
            @RequestParam("endRow") int endRow,
            @RequestParam("startColumn") int startColumn,
            @RequestParam("endColumn") int endColumn,
            @RequestParam(value = "sheet", required = false) String sheetName,
            @RequestParam(value = "sheetIndex", required = false, defaultValue = "0") int sheetIndex) {

        try{
            Sheet sheet = getSheet(file, sheetName, sheetIndex);

            List<SecondTestUserModel> users = XlsxConverter.extractObjectsFromTable(sheet, startRow, endRow, startColumn, endColumn, SecondTestUserModel.class);
            users.forEach(System.out::println);

            return ResponseEntity.ok().build();
        }catch (Exception e){
            e.printStackTrace();
        }
        return ResponseEntity.badRequest().build();
    }

    private static Sheet getSheet(MultipartFile file, String sheetName, int sheetIndex) throws IOException, InvalidFileTypeException, SheetNotFoundException {
        Workbook workbook;

        if(file.getOriginalFilename().endsWith(".xlsx")){
            workbook = new XSSFWorkbook(file.getInputStream());
        }else if(file.getOriginalFilename().endsWith(".xlsx")){
            workbook = new HSSFWorkbook(file.getInputStream());
        }else{
            throw new InvalidFileTypeException();
        }

        Sheet sheet;

        if(sheetName != null && !sheetName.equalsIgnoreCase("")){
            sheet = workbook.getSheet(sheetName);
        }else{
            sheet = workbook.getSheetAt(sheetIndex);
        }

        if(sheet == null) throw new SheetNotFoundException();

        return sheet;
    }
}
