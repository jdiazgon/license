package com.capgemini.license;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.ScrollPaneConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

import fr.opensagres.xdocreport.core.XDocReportException;

/**
 *
 */
public class UserView implements ActionListener {

    JButton openExcelBtn = new JButton("Open Excel");

    JButton generateDocBtn = new JButton("Open Doc");

    JPanel buttonPanel = new JPanel();

    JPanel masterPanel = new JPanel();

    File selectedExcelFile;

    File selectedDocFile;

    LicenseController licenseController;

    public UserView() {
        // Creating new frame
        JFrame frame = new JFrame("LicenseApp");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Setting listeners to buttons
        openExcelBtn.addActionListener(this);
        generateDocBtn.addActionListener(this);

        // Adding buttons to panel
        openExcelBtn.setBounds(120, 50, 140, 30);
        generateDocBtn.setBounds(120, 150, 140, 30);
        buttonPanel.setLayout(null);
        buttonPanel.add(openExcelBtn);
        buttonPanel.add(generateDocBtn);

        masterPanel.setLayout(new BorderLayout());
        masterPanel.add(buttonPanel, BorderLayout.CENTER);
        disableDocButton();

        // Setting Text area for console
        JTextArea textArea = new JTextArea(10, 30);
        textArea.setEditable(false);
        masterPanel.add(new JScrollPane(textArea, ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
            ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED), BorderLayout.SOUTH);

        // Setting frame settings
        frame.add(masterPanel);
        frame.setSize(400, 400);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);

        JTextAreaOutputStream out = new JTextAreaOutputStream(textArea);
        System.setOut(new PrintStream(out));
    }

    @Override
    public void actionPerformed(ActionEvent actionEvent) {
        String userDir = System.getProperty("user.home");
        String buttonName = actionEvent.getActionCommand();

        if (buttonName.equals("Open Excel")) {
            JFileChooser fileChooser = new JFileChooser(userDir + "/Desktop");
            FileNameExtensionFilter filter = new FileNameExtensionFilter("XLS files", "xls", "xlsx");
            fileChooser.setFileFilter(filter);

            int returnVal = fileChooser.showOpenDialog(masterPanel);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                System.out.println("You chose to open this file: " + fileChooser.getSelectedFile().getName());
                setSelectedExcelFile(fileChooser.getSelectedFile());
                licenseController.updateExcelModel(selectedExcelFile);
            }
        } else if (buttonName.equals("Open Doc")) {
            JFileChooser fileChooser = new JFileChooser(userDir + "/Desktop");
            FileNameExtensionFilter filter = new FileNameExtensionFilter("MS Word file(.docx)", "docx");
            fileChooser.setFileFilter(filter);
            int returnVal = fileChooser.showOpenDialog(masterPanel);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                System.out.println("You chose to open this file: " + fileChooser.getSelectedFile().getName());
                setSelectedDocFile(fileChooser.getSelectedFile());
                try {
                    licenseController.generateDocFile(selectedDocFile);
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (XDocReportException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
    }

    public File getSelectedExcelFile() {
        return selectedExcelFile;
    }

    public void setSelectedExcelFile(File selectedExcelFile) {
        this.selectedExcelFile = selectedExcelFile;
    }

    public File getSelectedDocFile() {
        return selectedDocFile;
    }

    public void setSelectedDocFile(File selectedDocFile) {
        this.selectedDocFile = selectedDocFile;
    }

    public LicenseController getLicenseController() {
        return licenseController;
    }

    public void setLicenseController(LicenseController licenseController) {
        this.licenseController = licenseController;
    }

    public void enableDocButton() {
        generateDocBtn.setEnabled(true);
    }

    public void disableDocButton() {
        generateDocBtn.setEnabled(false);
    }
}
