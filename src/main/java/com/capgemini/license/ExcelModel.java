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
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class that retrieves all the data from the excel
 */
public class ExcelModel {

    /**
     * Here is stated the column and row where the components version are stated
     */
    private static final String CELL_OF_COMPONENTS = "D1";

    private static final String PLUG_INS = "Plug-ins";

    private static final String MAIN_APPLICATIONS = "Main Applications";

    private static final String EMBEDDED_COMPONENTS = "Embedded Components";

    private static final String LANGUAGE = "Sprache:";

    File selectedExcelFile;

    private String language;

    private List<MainApplication> mainApplications;

    private List<EmbeddedComponent> embeddedComponents;

    private List<PlugIn> plugIns;

    private List<String> comments;

    private Workbook workbook = null;

    public ExcelModel() {
        mainApplications = new ArrayList<MainApplication>();
        embeddedComponents = new ArrayList<EmbeddedComponent>();
        plugIns = new ArrayList<PlugIn>();
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
            workbook = new XSSFWorkbook(excelFile);
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

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        // printList(mainApplications);
        // printList(embeddedComponents);
        // printList(plugIns);
        return true;
    }

    /**
     * Converts any cell value to string
     * @param currentCell
     */
    private String convertCellValueToString(Cell currentCell) {

        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            return currentCell.getStringCellValue();
        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
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
            return currentCell.getRichStringCellValue().getString();
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(currentCell)) {
                return currentCell.getDateCellValue().toString();
            } else {
                String returnValue = String.valueOf(currentCell.getNumericCellValue());
                return returnValue;
            }
        case BOOLEAN:
            return String.valueOf(currentCell.getBooleanCellValue());
        case ERROR:
            String errorMessage = "#REF error, code: " + currentCell.getErrorCellValue();
            return errorMessage;
        default:
            return "default formula cell";
        }
    }

    /**
     * Updates the Excel model by adding the information retrieved to the correct component
     *
     * @param stringCellValue
     *            Last string retrieved from the Excel
     * @param cellIterator
     *            The next cells of the row
     * @param iterator
     *            The next rows of the file
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
     * Traverse through the Excel file for finding all the plug-ins.
     * @param rowIterator
     *            The next rows of the file
     */
    private void updatePlugIns(Iterator<Row> rowIterator) {
        String stringCellValue;
        Iterator<Cell> cellIterator;
        plugIns = new ArrayList<PlugIn>();

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();

            int cellCounter = 0;
            String plugInName = "";
            String plugInVersion = "";
            String plugInDescription = "";

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (isNotBlank(stringCellValue)) {
                    // Setting name of the plugIn
                    if (cellCounter == 0) {
                        plugInName = stringCellValue;
                    }
                    // Setting version of the plugIn
                    if (cellCounter == 1) {
                        plugInVersion = stringCellValue;
                    }
                    // Setting description of the plugIn and creating a new PlugIn object
                    if (cellCounter == 2) {
                        plugInDescription = stringCellValue;
                        PlugIn plugIn = new PlugIn(plugInName, plugInVersion, plugInDescription);
                        plugIns.add(plugIn);
                    }
                    cellCounter++;
                }
            }
        }
    }

    /**
     * Traverse through the Excel file for finding all the embedded components until a plug-in is found.
     * @param stringCellValue
     *            Last string retrieved from the Excel
     * @param rowIterator
     *            The next rows of the file
     * @return Last string retrieved from the Excel
     */
    private String updateEmbeddedComponents(Iterator<Row> rowIterator) {
        String stringCellValue = "BLANK";
        Iterator<Cell> cellIterator;
        embeddedComponents = new ArrayList<EmbeddedComponent>();

        while (rowIterator.hasNext()) {
            Boolean breakWhile = false;
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();

            int cellCounter = 0;
            String embeddedComponentName = "";
            String embeddedComponentVersion = "";
            String embeddedComponentDescription = "";

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (stringCellValue.equals(PLUG_INS)) {
                    breakWhile = true;
                    break;
                }
                if (isNotBlank(stringCellValue)) {
                    // Setting name of the embedded component
                    if (cellCounter == 0) {
                        embeddedComponentName = stringCellValue;
                    }
                    // Setting version of the embedded component
                    if (cellCounter == 1) {
                        embeddedComponentVersion = stringCellValue;
                    }
                    // Setting description of the embedded component and creating a new EmbeddedComponent
                    // object
                    if (cellCounter == 2) {
                        embeddedComponentDescription = stringCellValue;
                        EmbeddedComponent embeddedComponent = new EmbeddedComponent(embeddedComponentName,
                            embeddedComponentVersion, embeddedComponentDescription);
                        embeddedComponents.add(embeddedComponent);
                    }
                    cellCounter++;
                }
            }
            if (breakWhile) {
                break;
            }
        }
        return stringCellValue;
    }

    /**
     * Traverse through the Excel file for finding all the main applications until an embedded components is
     * found.
     * @param stringCellValue
     *            Last string retrieved from the Excel
     * @param rowIterator
     *            The next rows of the file
     * @return Last string retrieved from the Excel
     */
    private String updateMainApplications(Iterator<Row> rowIterator) {
        String stringCellValue = "BLANK";
        Iterator<Cell> cellIterator;
        mainApplications = new ArrayList<MainApplication>();

        while (rowIterator.hasNext()) {
            Boolean breakWhile = false;
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();

            int cellCounter = 0;
            String mainApplicationName = "";
            String mainApplicationVersion = "";
            String mainApplicationDescription = "";

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (stringCellValue.equals(EMBEDDED_COMPONENTS)) {
                    breakWhile = true;
                    break;
                }
                if (isNotBlank(stringCellValue)) {
                    // Setting name of the main application
                    if (cellCounter == 0) {
                        mainApplicationName = stringCellValue;
                    }
                    // Setting version of the main application
                    if (cellCounter == 1) {
                        mainApplicationVersion = stringCellValue;
                    }
                    // Setting description of the main application and creating a new MainApplication object
                    if (cellCounter == 2) {
                        mainApplicationDescription = stringCellValue;
                        MainApplication mainApplication = new MainApplication(mainApplicationName,
                            mainApplicationVersion, mainApplicationDescription);
                        mainApplications.add(mainApplication);
                    }
                    cellCounter++;
                }
            }
            if (breakWhile) {
                break;
            }
        }
        return stringCellValue;
    }

    /**
     * Tries to find a sheet with his name for getting all the component's data
     * @param name
     *            name of the plug-in
     * @param version
     *            version of the plug-in
     */
    public List<Component> getComponents(String name, String version) {
        Sheet datatypeSheet = workbook.getSheet(name);
        if (datatypeSheet != null) {
            Iterator<Row> iterator = datatypeSheet.iterator();
            List<Component> components = new ArrayList<Component>();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    String cellValue = convertCellValueToString(currentCell);
                    components = updateComponentModel(cellValue, cellIterator, iterator, version, name);
                    if (components.size() >= 1) {
                        return components;
                    }
                }

            }
            if (components.size() >= 1) {
                return components;
            }
        }

        return null;
    }

    /**
     * Creates a new list of components and checks if the plug-in version contains components
     * @param cellValue
     *            Last string retrieved from the Excel
     * @param cellIterator
     *            The next cells of the row
     * @param iterator
     *            The next rows of the file
     * @param version
     *            Version of the plug-in
     * @param lastRowNumber
     * @return A list of components
     */
    private List<Component> updateComponentModel(String cellValue, Iterator<Cell> cellIterator, Iterator<Row> iterator,
        String version, String sheetName) {
        List<Component> components = new ArrayList<Component>();
        if (cellValue.equals(version)) {
            // Component found, lets get the table data
            components = getComponentData(iterator, version, sheetName);
        }
        return components;
    }

    /**
     * Traverse through the components and retrieve the correct information.
     * @param rowIterator
     * @param version
     *            Version of the plug-in
     * @param sheetName
     *            Used for getting the current sheet
     * @param iterator
     *            The next rows of the file
     * @return A list of components
     */
    private List<Component> getComponentData(Iterator<Row> rowIterator, String version, String sheetName) {

        String stringCellValue = "BLANK";
        Iterator<Cell> cellIterator;
        List<Component> components = new ArrayList<Component>();
        comments = new ArrayList<String>();
        // Let's move to the position where the components are stated
        Sheet sheet = workbook.getSheet(sheetName);
        int firstRowToGet = getRowNumberOfCurrentComponentVersion(version, sheet);

        for (int i = firstRowToGet; i <= sheet.getLastRowNum(); i++) {
            Boolean breakWhile = false;
            Row currentRow = sheet.getRow(i);
            if (currentRow != null) {
                cellIterator = currentRow.iterator();

                int cellCounter = 0;
                String componentName = "";
                String componentVersion = "";
                String componentLicense = "";

                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    stringCellValue = convertCellValueToString(currentCell);
                    if (isVersion(stringCellValue)) {
                        breakWhile = true;
                        break;
                    }

                    if (isNotBlank(stringCellValue)) {
                        // Setting name of the component
                        if (cellCounter == 0) {
                            String checkComment = stringCellValue.substring(0, 3);
                            if (isComment(checkComment)) {
                                comments.add(stringCellValue.substring(3));
                            } else {
                                componentName = stringCellValue;
                            }
                        }
                        // Setting version of the embedded component
                        if (cellCounter == 1) {
                            componentVersion = stringCellValue;
                        }
                        // Setting description of the embedded component and creating a new EmbeddedComponent
                        // object
                        if (cellCounter == 2) {
                            componentLicense = stringCellValue;
                            Component component = new Component(componentName, componentVersion, componentLicense);
                            components.add(component);
                        }
                        cellCounter++;
                    }
                }
            }
            if (breakWhile) {
                break;
            }
        }
        return components;
    }

    /**
     * Tries to find the row number where the current version components are stated
     * @param version
     *            version of the current plugIn
     * @param sheet
     *            The sheet of the current plugIn
     */
    private Integer getRowNumberOfCurrentComponentVersion(String version, Sheet sheet) {
        // Column where the components should be stated
        CellReference cr = new CellReference(CELL_OF_COMPONENTS);
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow != null) {
                Cell cell = currentRow.getCell(cr.getCol());
                if (cell != null) {
                    // Found column and there is value in the cell.
                    String cellValue = cell.getStringCellValue();
                    if (cellValue.equals(version)) {
                        return i + 1;
                    }
                }
            }
        }
        return sheet.getLastRowNum();
    }

    /**
     * @param version
     * @param stringCellValue
     * @return
     */
    private boolean isNotCurrentVersion(String version, String stringCellValue) {
        return !stringCellValue.equals(version);
    }

    /**
     * @param stringCellValue
     * @return
     */
    private boolean isVersion(String stringCellValue) {
        return stringCellValue.subSequence(0, 3).toString().matches("^[v][0-9].");
    }

    /**
     * Checks if the text found is a comment and if it is the one for the selected language
     * @param checkComment
     * @return true: If it's a comment for our language
     */
    private boolean isComment(String checkComment) {
        return checkComment.equals(getLanguage() + ":");
    }

    /**
     * Checks if string value is not "BLANK"
     * @param cellValue
     *            Last string retrieved from the Excel
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

    public List<MainApplication> getMainApplications() {
        return mainApplications;
    }

    public List<String> getComments() {
        return comments;
    }

    public void setMainApplications(List<MainApplication> mainApplications) {
        this.mainApplications = mainApplications;
    }

    public List<EmbeddedComponent> getEmbeddedComponents() {
        return embeddedComponents;
    }

    public void setEmbeddedComponents(List<EmbeddedComponent> embeddedComponents) {
        this.embeddedComponents = embeddedComponents;
    }

    public List<PlugIn> getPlugIns() {
        return plugIns;
    }

    public void setPlugIns(List<PlugIn> plugIns) {
        this.plugIns = plugIns;
    }

}
