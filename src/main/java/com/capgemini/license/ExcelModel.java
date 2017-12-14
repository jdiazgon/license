package com.capgemini.license;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 */
public class ExcelModel {

    /**
     *
     */
    private static final String MAIN_APPLICATIONS = "Main Applications";

    /**
     *
     */
    private static final String LANGUAGE = "Sprache:";

    File selectedExcelFile;

    private String language;

    private List<String> mainApplications;

    private List<String> embeddedComponents;

    private List<String> plugIns;

    public ExcelModel() {
        mainApplications = new ArrayList<String>();
        embeddedComponents = new ArrayList<String>();
        plugIns = new ArrayList<String>();
    }

    public File getSelectedExcelFile() {
        return selectedExcelFile;
    }

    public void setSelectedExcelFile(File selectedExcelFile) {
        this.selectedExcelFile = selectedExcelFile;
    }

    /**
     *
     */
    public void readExcelData() {
        try {

            FileInputStream excelFile = new FileInputStream(selectedExcelFile);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    // getCellTypeEnum shown as deprecated for version 3.15
                    // getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                    String cellValue = convertCellValueToString(currentCell);
                    updateModel(cellValue, cellIterator, iterator);
                }
                System.out.println();

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * @param currentCell
     */
    private String convertCellValueToString(Cell currentCell) {
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            System.out.print(currentCell.getStringCellValue() + "--");
            return currentCell.getStringCellValue();
        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
            System.out.print(currentCell.getNumericCellValue() + "--");
            String returnValue = String.valueOf(currentCell.getNumericCellValue());
            return returnValue;
        }
        return "BLANK";
    }

    // updateModel(currentCell.getStringCellValue(), cellIterator, iterator);
    /**
     * @param stringCellValue
     * @param cellIterator
     * @param iterator
     */
    private void updateModel(String stringCellValue, Iterator<Cell> cellIterator, Iterator<Row> iterator) {

        if (stringCellValue.equals(LANGUAGE)) {
            Cell nextCell = cellIterator.next();
            setLanguage(nextCell.getStringCellValue());
        } else if (stringCellValue.equals(MAIN_APPLICATIONS)) {
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    String cellValue = convertCellValueToString(currentCell);
                }
            }
        }
    }

    public String getLanguage() {
        return language;
    }

    public void setLanguage(String language) {
        this.language = language;
    }

}
