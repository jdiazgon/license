package com.capgemini.license;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.net.URL;

import javax.swing.ImageIcon;
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

    JButton openDocBtn = new JButton("Open Doc");

    JButton generateDocBtn = new JButton("Generate");

    JPanel buttonPanel = new JPanel();

    JPanel masterPanel = new JPanel();

    File selectedExcelFile;

    File selectedDocFile;

    File selectedGenerationFolder;

    LicenseController licenseController;

    public UserView() {
        // Creating new frame
        JFrame frame = new JFrame("LicenseApp");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Set a nice icon to the program
        URL iconURL = getClass().getResource("/appIcon.png");
        if (iconURL != null) {
            // iconURL is null when not found
            ImageIcon icon = new ImageIcon(iconURL);
            frame.setIconImage(icon.getImage());
        }

        // Setting listeners to buttons
        openExcelBtn.addActionListener(this);
        openDocBtn.addActionListener(this);
        generateDocBtn.addActionListener(this);

        // Adding buttons to panel
        openExcelBtn.setPreferredSize(new Dimension(120, 30));
        openDocBtn.setPreferredSize(new Dimension(120, 30));
        generateDocBtn.setPreferredSize(new Dimension(120, 30));

        buttonPanel.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.weighty = 0.5;
        buttonPanel.add(openExcelBtn, gbc);

        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.weighty = 0.5;
        buttonPanel.add(openDocBtn, gbc);

        gbc.gridx = 0;
        gbc.gridy = 6;
        gbc.weighty = 0.5;
        buttonPanel.add(generateDocBtn, gbc);

        masterPanel.setLayout(new BorderLayout());
        masterPanel.add(buttonPanel, BorderLayout.CENTER);
        disableDocButton();
        disableGenerateButton();

        // Setting Text area for console
        JTextArea textArea = new JTextArea(10, 30);
        textArea.setEditable(false);
        masterPanel.add(new JScrollPane(textArea, ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
            ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED), BorderLayout.SOUTH);

        // Setting frame settings
        frame.add(masterPanel);
        frame.setSize(600, 400);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);

        JTextAreaOutputStream out = new JTextAreaOutputStream(textArea);
        System.setOut(new PrintStream(out));
        System.setErr(new PrintStream(out));
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
                System.out.println("-------------------------------------------------------------");
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
                enableGenerateButton();
            }
        } else if (buttonName.equals("Generate")) {
            JFileChooser folderChooser = new JFileChooser(userDir + "/Desktop");
            folderChooser.setDialogTitle("Select a location for generation");

            // Default generated file name
            folderChooser.setSelectedFile(new File("Softwarelizenzvertrag CobiGen.docx"));

            if (folderChooser.showSaveDialog(masterPanel) == JFileChooser.APPROVE_OPTION) {
                System.out.println("");
                File chosenFile = folderChooser.getSelectedFile();
                // Checks if the file is a doc, and then saves it
                setSelectedGenerationFolder(checkFileForExtension(chosenFile));
                System.out.println("You chose to store it on : " + folderChooser.getSelectedFile());
                try {
                    licenseController.generateDocFile(selectedDocFile);
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                } catch (XDocReportException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            } else {
                System.out.println("You did not select a folder...");
            }
        }
    }

    /**
     * This method checks if the extension of the parameter is ".docx'. If not, then a new File is created
     * with this extension, and returned. If yes, then the parameter itself is returned.
     *
     * @param f
     *            File, the file in question
     * @return a file with extension '.docx'. This can be the parameter itself.
     */
    private File checkFileForExtension(File f) {
        if (f.getName().toLowerCase().endsWith(".docx")) {
            return f;
        } else {
            return new File(f.getParent(), f.getName() + ".docx");
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

    public File getSelectedGenerationFolder() {
        return selectedGenerationFolder;
    }

    public void setSelectedGenerationFolder(File selectedGenerationFolder) {
        this.selectedGenerationFolder = selectedGenerationFolder;
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
        openDocBtn.setEnabled(true);
    }

    public void disableDocButton() {
        openDocBtn.setEnabled(false);
    }

    public void enableGenerateButton() {
        generateDocBtn.setEnabled(true);
    }

    public void disableGenerateButton() {
        generateDocBtn.setEnabled(false);
    }
}
