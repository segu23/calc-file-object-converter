package com.github.segu23.xlsxconverter.controller;

import com.github.segu23.xlsxconverter.converter.CalcFileConverter;
import com.github.segu23.xlsxconverter.model.CriteriaModel;
import com.github.segu23.xlsxconverter.model.DGNPassCriteriaIndexModel;
import com.github.segu23.xlsxconverter.model.DGNTriggerModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
@RequestMapping("/dgn")
public class DGNController {

    static {
        CalcFileConverter.addDataTypeConversion(CriteriaModel.class, data -> {
            String[] dataSplit = data.split(" = ", 2);

            if (dataSplit.length < 2) {
                return new CriteriaModel(null, data);
            } else {
                String criteriaName = dataSplit[0];
                String criteriaValue = dataSplit[1];

                return new CriteriaModel(criteriaName, criteriaValue);
            }
        });
    }

    @PostMapping("/upload")
    public void uploadDocument(@RequestParam("file") MultipartFile file) {
        try {
            Workbook workbook = CalcFileConverter.getWorkbook(file);
            // Triggers
            Sheet triggersSheet = CalcFileConverter.getSheetByName(workbook, "Triggers");

            int triggersEndRow = CalcFileConverter.getLastRow(triggersSheet, 0, 0);
            int triggersEndColumn = CalcFileConverter.getLastColumn(triggersSheet, 0, 0);

            List<DGNTriggerModel> triggers = CalcFileConverter.extractObjectsFromTable(triggersSheet, 0, triggersEndRow, 0, triggersEndColumn, DGNTriggerModel.class);
            triggers.forEach(System.out::println);

            // Pass criteria indexes
            Sheet passCriteriaIndexesSheet = CalcFileConverter.getSheetByName(workbook, "Pass criteria Indexes");

            int firstPassCriteriaIndexesEndRow = CalcFileConverter.getLastRow(passCriteriaIndexesSheet, 0, 0);
            int firstPassCriteriaIndexesEndColumn = CalcFileConverter.getLastColumn(passCriteriaIndexesSheet, 0, 0);

            List<DGNPassCriteriaIndexModel> firstPassCriteriaIndexes = CalcFileConverter.extractObjectsFromTable(
                    passCriteriaIndexesSheet,
                    0,
                    firstPassCriteriaIndexesEndRow,
                    0,
                    firstPassCriteriaIndexesEndColumn,
                    DGNPassCriteriaIndexModel.class,
                    new String[]{"Index", "Criteria"}
            );

            firstPassCriteriaIndexes.forEach(System.out::println);

            int secondPassCriteriaIndexesEndRow = CalcFileConverter.getLastRow(passCriteriaIndexesSheet, 0, 3);
            int secondPassCriteriaIndexesEndColumn = CalcFileConverter.getLastColumn(passCriteriaIndexesSheet, 0, 3);

            List<DGNPassCriteriaIndexModel> secondPassCriteriaIndexes = CalcFileConverter.extractObjectsFromTable(
                    passCriteriaIndexesSheet,
                    0,
                    secondPassCriteriaIndexesEndRow,
                    3,
                    secondPassCriteriaIndexesEndColumn,
                    DGNPassCriteriaIndexModel.class,
                    new String[]{"Index", "Criteria"}
            );

            secondPassCriteriaIndexes.forEach(System.out::println);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
