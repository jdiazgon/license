package com.capgemini.license;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;

/**
 * In charge of starting the excel reading and updating the User Interface
 */
public class LicenseController {
    private ExcelModel excelModel;

    private UserView userView;

    // Java model context for generating files
    private IContext context;

    // List for every component of every plugin
    private List<List<Component>> pluginsComponents;

    /**
     * MVC: Controller module
     * @param excelModel
     *            Contains the data from the Excel File
     * @param userView
     *            User interface for opening files
     */
    public LicenseController(ExcelModel excelModel, UserView userView) {
        this.excelModel = excelModel;
        this.userView = userView;
    }

    /**
     * Starts the excel reading of the file and outputs the quantity of elements retrieved.
     * @param selectedExcelFile
     *            Excel file to be read
     */
    public void updateExcelModel(File selectedExcelFile) {
        excelModel.setSelectedExcelFile(selectedExcelFile);

        if (excelModel.readExcelData()) {
            System.out.println();
            System.out.println("Main Applications contains: " + excelModel.getMainApplications().size() + " elements.");
            System.out
                .println("Embedded components contains: " + excelModel.getEmbeddedComponents().size() + " elements.");
            System.out.println("Plugins contains: " + excelModel.getPlugIns().size() + " elements.");
            System.out.println();

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

        context = report.createContext();
        setContextData();

        String pathFile = selectedDocFile.getParent();

        // Generate report by merging Java model with the DOC
        File outputFile = new File(pathFile + "\\OutputDoc.docx");
        selectedDocFile.getCanonicalPath();
        OutputStream out = new FileOutputStream(outputFile);
        if (outputFile.canWrite()) {
            // write access
            report.process(context, out);
            // Closing OutputStream
            out.flush();
            out.close();
            System.out.println("File succesfully generated");
        } else {
            // no write access
            System.out.println("No write acces");
        }
    }

    /**
     * Sets all the data retrieved from the Excel to the IContext
     */
    private void setContextData() {
        context.put("name", "world");
        context.put("mainApplications", excelModel.getMainApplications());
        context.put("embeddedComponents", excelModel.getEmbeddedComponents());
        context.put("plugIns", excelModel.getPlugIns());

        List<MainApplication> mainApplications = excelModel.getMainApplications();
        pluginsComponents = new ArrayList<List<Component>>();

        for (MainApplication mainApplication : mainApplications) {
            List<Component> components = new ArrayList<Component>();
            components = excelModel.getComponents(mainApplication.getName(), mainApplication.getVersion());
            pluginsComponents.add(components);
        }
        context.put("mainApplicationsComponentsList", pluginsComponents);

        List<EmbeddedComponent> embeddedComponents = excelModel.getEmbeddedComponents();
        pluginsComponents = new ArrayList<List<Component>>();
        for (EmbeddedComponent embeddedComponent : embeddedComponents) {
            List<Component> components = new ArrayList<Component>();
            components = excelModel.getComponents(embeddedComponent.getName(), embeddedComponent.getVersion());
            pluginsComponents.add(components);
        }
        context.put("embeddedComponentsComponentsList", pluginsComponents);

        List<PlugIn> plugIns = excelModel.getPlugIns();
        pluginsComponents = new ArrayList<List<Component>>();
        for (PlugIn plugIn : plugIns) {
            List<Component> components = new ArrayList<Component>();
            components = excelModel.getComponents(plugIn.getName(), plugIn.getVersion());
            if (components != null) {
                pluginsComponents.add(components);
            }
        }
        context.put("plugInsComponentsList", pluginsComponents);

    }
}
