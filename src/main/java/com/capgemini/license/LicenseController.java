package com.capgemini.license;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;

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
        if (excelModel.readExcelData()) {
            enableDocButton();
        }
    }

    public void enableDocButton() {
        userView.enableDocButton();
    }

    public void disableDocButton() {
        userView.disableDocButton();
    }

    /**
     * Method for the generation of the Doc file. It opens a file, sets the template engine, creates the Data
     * model and generates the document.
     * @param selectedDocFile
     * @throws XDocReportException
     * @throws IOException
     */
    public void generateDocFile(File selectedDocFile) throws IOException, XDocReportException {
        // Load DOC file and set FreeMarker template engine and cache it to the registry
        InputStream in = new FileInputStream(selectedDocFile);
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(in, TemplateEngineKind.Freemarker);

        // Create Java model context
        IContext context = report.createContext();
        context.put("name", "world");
        context.put("mainApplications", excelModel.getMainApplications());
        context.put("embeddedComponents", excelModel.getEmbeddedComponents());
        context.put("plugIns", excelModel.getPlugIns());

        String pathFile = selectedDocFile.getParent();

        // Generate report by merging Java model with the DOC
        File outputFile = new File(pathFile + "\\OutputDoc.docx");
        selectedDocFile.getCanonicalPath();
        OutputStream out = new FileOutputStream(outputFile);
        if (outputFile.canWrite()) {
            // write access
            report.process(context, out);
        } else {
            // no write access
            System.out.println("No write acces");
        }
    }
}
