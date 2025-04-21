package com.qa.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelComparator {

    public static void main(String[] args) {
        String filePath1 = "path_to_first_file.xlsx";
        String filePath2 = "path_to_second_file.xlsx";

        try {
            boolean areFilesIdentical = compareExcelFiles(filePath1, filePath2);
            if (areFilesIdentical) {
                System.out.println("The Excel files are identical.");
            } else {
                System.out.println("The Excel files have differences.");
            }
        } catch (IOException e) {
            System.err.println("Error comparing Excel files: " + e.getMessage());
        }
    }

    public static boolean compareExcelFiles(String filePath1, String filePath2) throws IOException {
        try (FileInputStream fis1 = new FileInputStream(new File(filePath1));
             FileInputStream fis2 = new FileInputStream(new File(filePath2));
             Workbook workbook1 = new XSSFWorkbook(fis1);
             Workbook workbook2 = new XSSFWorkbook(fis2)) {

            int sheetCount1 = workbook1.getNumberOfSheets();
            int sheetCount2 = workbook2.getNumberOfSheets();

            if (sheetCount1 != sheetCount2) {
                System.out.println("The number of sheets is different.");
                return false;
            }

            for (int i = 0; i < sheetCount1; i++) {
                Sheet sheet1 = workbook1.getSheetAt(i);
                Sheet sheet2 = workbook2.getSheetAt(i);

                if (!compareSheets(sheet1, sheet2)) {
                    System.out.println("Difference found in sheet: " + sheet1.getSheetName());
                    return false;
                }
            }
        }
        return true;
    }

    private static boolean compareSheets(Sheet sheet1, Sheet sheet2) {
        int lastRowNum1 = sheet1.getLastRowNum();
        int lastRowNum2 = sheet2.getLastRowNum();

        if (lastRowNum1 != lastRowNum2) {
            System.out.println("The number of rows is different in sheet: " + sheet1.getSheetName());
            return false;
        }

        for (int i = 0; i <= lastRowNum1; i++) {
            Row row1 = sheet1.getRow(i);
            Row row2 = sheet2.getRow(i);

            if (!compareRows(row1, row2)) {
                System.out.println("Difference found in cell: " + i);
                return false;
            }
        }

        return true;
    }

    private static boolean compareRows(Row row1, Row row2) {
        if ((row1 == null && row2 != null) || (row1 != null && row2 == null)) {
            return false;
        }
        if (row1 == null) {
            return true; // Both rows are null
        }

        int lastCellNum1 = row1.getLastCellNum();
        int lastCellNum2 = row2.getLastCellNum();

        if (lastCellNum1 != lastCellNum2) {
            return false;
        }

        for (int i = 0; i < lastCellNum1; i++) {
            Cell cell1 = row1.getCell(i);
            Cell cell2 = row2.getCell(i);

            if (!compareCells(cell1, cell2)) {
                System.out.println("Difference found in cell: " + i);
                return false;
            }
        }

        return true;
    }

    private static boolean compareCells(Cell cell1, Cell cell2) {
        if ((cell1 == null && cell2 != null) || (cell1 != null && cell2 == null)) {
            return false;
        }
        if (cell1 == null) {
            return true; // Both cells are null
        }

        if (cell1.getCellType() != cell2.getCellType()) {
            return false;
        }

        switch (cell1.getCellType()) {
            case STRING:
                return cell1.getStringCellValue().equals(cell2.getStringCellValue());
            case NUMERIC:
                return Double.compare(cell1.getNumericCellValue(), cell2.getNumericCellValue()) == 0;
            case BOOLEAN:
                return cell1.getBooleanCellValue() == cell2.getBooleanCellValue();
            case FORMULA:
                return cell1.getCellFormula().equals(cell2.getCellFormula());
            case BLANK:
                return true; // Both are blank
            default:
                return false;
        }
    }
}
