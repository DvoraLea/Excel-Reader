package com.example.demo;

import org.apache.poi.ss.usermodel.*;

public class SheetProcessor {

    public static FormulaEvaluator getFormulaEvaluator(Sheet sheet) {
        return sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
    }
}
