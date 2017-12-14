package com.capgemini.license;

/**
 * Hello world!
 *
 */
public class LicenseApp {
    public static void main(String[] args) {
        UserView userView = new UserView();
        ExcelModel excelModel = new ExcelModel();
        LicenseController licenseController = new LicenseController(excelModel, userView);

        userView.setLicenseController(licenseController);
    }
}
