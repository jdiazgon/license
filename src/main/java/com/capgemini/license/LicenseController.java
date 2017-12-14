package com.capgemini.license;

import java.io.File;

/**
 *
 */
public class LicenseController {
    private ExcelModel excelModel;

    private UserView userView;

    /**
     * MVC: Controller module
     * @param excelModel
     *            Contains the data from the Excel File
     * @param userView
     *            User interface for opening files
     */
    public LicenseController(ExcelModel excelModel, UserView userView) {
        super();
        this.excelModel = excelModel;
        this.userView = userView;
    }

    public void updateExcelModel(File selectedExcelFile) {
        excelModel.setSelectedExcelFile(selectedExcelFile);
        excelModel.readExcelData();
    }
}
