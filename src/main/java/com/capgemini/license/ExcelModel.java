package com.capgemini.license;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 */
public class ExcelModel {

    private static final String PLUG_INS = "Plug-ins";

    private static final String MAIN_APPLICATIONS = "Main Applications";

    private static final String EMBEDDED_COMPONENTS = "Embedded Components";

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
     * Reads every content of an excel file
     */
    public Boolean readExcelData() {
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
                    String cellValue = convertCellValueToString(currentCell);
                    updateModel(cellValue, cellIterator, iterator);
                }
                System.out.println();

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        printList(mainApplications);
        printList(embeddedComponents);
        printList(plugIns);
        return true;
    }

    /**
     * Converts any cell value to string
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
        } else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
            return convertFormulaValueToString(currentCell);
        }
        return "BLANK";
    }

    /**
     * Checks which kind of formula value is retrieved and converts it to a string
     * @param currentCell
     */
    private String convertFormulaValueToString(Cell currentCell) {
        switch (currentCell.getCachedFormulaResultTypeEnum()) {
        case STRING:
            System.out.println(currentCell.getRichStringCellValue().getString());
            return currentCell.getRichStringCellValue().getString();
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(currentCell)) {
                System.out.println(currentCell.getDateCellValue());
                return currentCell.getDateCellValue().toString();
            } else {
                System.out.println(currentCell.getNumericCellValue());
                String returnValue = String.valueOf(currentCell.getNumericCellValue());
                return returnValue;
            }
        case BOOLEAN:
            System.out.println(currentCell.getBooleanCellValue());
            return String.valueOf(currentCell.getBooleanCellValue());
        case ERROR:
            System.out.println(currentCell.getErrorCellValue());
            return String.valueOf(currentCell.getErrorCellValue());
        default:
            System.out.println("default formula cell"); // should never occur
            return "default formula cell";
        }
    }

    /**
     * @param stringCellValue
     * @param cellIterator
     * @param iterator
     */
    private void updateModel(String stringCellValue, Iterator<Cell> cellIterator, Iterator<Row> iterator) {

        if (stringCellValue.equals(LANGUAGE)) {
            Cell nextCell = cellIterator.next();
            setLanguage(nextCell.getStringCellValue());
            return;
        } else if (stringCellValue.equals(MAIN_APPLICATIONS)) {
            stringCellValue = updateMainApplications(iterator);
        }
        if (stringCellValue.equals(EMBEDDED_COMPONENTS)) {
            stringCellValue = updateEmbeddedComponents(iterator);
        }
        if (stringCellValue.equals(PLUG_INS)) {
            updatePlugIns(iterator);
        }
    }

    /**
     * @param rowIterator
     */
    private void updatePlugIns(Iterator<Row> rowIterator) {
        String stringCellValue;
        Iterator<Cell> cellIterator;
        while (rowIterator.hasNext()) {
            Boolean breakWhile = false;
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (stringCellValue.equals(PLUG_INS)) {
                    breakWhile = true;
                    break;
                }
                if (isNotBlank(stringCellValue)) {
                    plugIns.add(stringCellValue);
                }
            }
            if (breakWhile) {
                break;
            }
        }
    }

    /**
     * @param stringCellValue
     * @param rowIterator
     * @return
     */
    private String updateEmbeddedComponents(Iterator<Row> rowIterator) {
        String stringCellValue = "BLANK";
        Iterator<Cell> cellIterator;
        while (rowIterator.hasNext()) {
            Boolean breakWhile = false;
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (stringCellValue.equals(PLUG_INS)) {
                    breakWhile = true;
                    break;
                }
                if (isNotBlank(stringCellValue)) {
                    embeddedComponents.add(stringCellValue);
                }
            }
            if (breakWhile) {
                break;
            }
        }
        return stringCellValue;
    }

    /**
     * @param stringCellValue
     * @param rowIterator
     * @return
     */
    private String updateMainApplications(Iterator<Row> rowIterator) {
        String stringCellValue = "BLANK";
        Iterator<Cell> cellIterator;
        while (rowIterator.hasNext()) {
            Boolean breakWhile = false;
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (stringCellValue.equals(EMBEDDED_COMPONENTS)) {
                    breakWhile = true;
                    break;
                }
                if (isNotBlank(stringCellValue)) {
                    mainApplications.add(stringCellValue);
                }
            }
            if (breakWhile) {
                break;
            }
        }
        return stringCellValue;
    }

    /**
     * @param cellValue
     * @return
     */
    private boolean isNotBlank(String cellValue) {
        return !cellValue.equals("BLANK");
    }

    public String getLanguage() {
        return language;
    }

    public void setLanguage(String language) {
        this.language = language;
    }

    public void printList(List<String> list) {
        System.out.println("List: ");
        System.out.println(Arrays.toString(list.toArray()));
    }

}
