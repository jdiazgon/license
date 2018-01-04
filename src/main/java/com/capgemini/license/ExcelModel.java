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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class that retrieves all the data from the excel
 */
public class ExcelModel {

    private static final String PLUG_INS = "Plug-ins";

    private static final String MAIN_APPLICATIONS = "Main Applications";

    private static final String EMBEDDED_COMPONENTS = "Embedded Components";

    private static final String LANGUAGE = "Sprache:";

    File selectedExcelFile;

    private String language;

    private List<MainApplication> mainApplications;

    private List<EmbeddedComponent> embeddedComponents;

    private List<PlugIn> plugIns;

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

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                if (isNotBlank(stringCellValue)) {
                    // Setting name of the plugIn
                    if (cellCounter == 0) {
                        plugInName = stringCellValue;
                    }
                    // Setting version of the plugIn and creating a new PlugIn object
                    if (cellCounter == 1) {
                        plugInVersion = stringCellValue;
                        PlugIn plugIn = new PlugIn(plugInName, plugInVersion);
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
                    components = updateComponentModel(cellValue, cellIterator, iterator, version);
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
     * @return A list of components
     */
    private List<Component> updateComponentModel(String cellValue, Iterator<Cell> cellIterator, Iterator<Row> iterator,
        String version) {
        List<Component> components = new ArrayList<Component>();
        if (cellValue.equals(version)) {
            // Component found, lets get the table data
            components = getComponentData(iterator);
        }
        return components;
    }

    /**
     * Traverse through the components and retrieve the correct information.
     * @param iterator
     *            The next rows of the file
     * @return A list of components
     */
    private List<Component> getComponentData(Iterator<Row> rowIterator) {

        String stringCellValue = "BLANK";
        Iterator<Cell> cellIterator;
        List<Component> components = new ArrayList<Component>();

        while (rowIterator.hasNext()) {
            Boolean breakWhile = false;
            Row currentRow = rowIterator.next();
            cellIterator = currentRow.iterator();

            int cellCounter = 0;
            String componentName = "";
            String componentVersion = "";
            String componentLicense = "";

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                stringCellValue = convertCellValueToString(currentCell);
                // TODO: Is there a better logic? This checks if the string is a version (E.g. v.4.0)
                if (stringCellValue.subSequence(0, 3).toString().matches("^[v][0-9].")) {
                    breakWhile = true;
                    break;
                }
                if (isNotBlank(stringCellValue)) {
                    // Setting name of the component
                    if (cellCounter == 0) {
                        componentName = stringCellValue;
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
            if (breakWhile) {
                break;
            }
        }
        return components;
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
