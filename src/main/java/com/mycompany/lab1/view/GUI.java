package com.mycompany.lab1.view;

import com.mycompany.lab1.controller.IController;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;

public class GUI extends JFrame {

    private JButton importButton;
    private JButton exportButton;
    private JButton exitButton;
    private IController controller;

    public GUI(IController controller) {
        super("Лабораторная работа 1");
        this.controller = controller;
        setLayout(new GridLayout(3, 1));

        importButton = new JButton("Выбрать .xlsx файл для импорта");
        add(importButton);

        importButton.addActionListener((ActionEvent e) -> {
            JFileChooser fileChooser = new JFileChooser();
            File defaultDirectory = new File("C:\\Users\\nsoko\\OneDrive\\Desktop\\for_lab1");
            fileChooser.setCurrentDirectory(defaultDirectory);
            int result = fileChooser.showOpenDialog(null);

            if (result == JFileChooser.APPROVE_OPTION) {
                String filePath = fileChooser.getSelectedFile().getAbsolutePath();
                try {
                    controller.controllerImportData(filePath);
                    setExportButtonEnabled(true);
                    JOptionPane.showMessageDialog(null, "Файл успешно импортирован!");
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Ошибка при импорте файла: " + ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        exportButton = new JButton("Создать .xlsx файл для рассчитанных показателей");
        exportButton.setEnabled(false);
        add(exportButton);

        exportButton.addActionListener((ActionEvent e) -> {
            if (controller.getStorage() == null) {
                JOptionPane.showMessageDialog(null, "Сначала импортируйте файл!", "Ошибка", JOptionPane.ERROR_MESSAGE);
                return;
            }

            String[] sheetNames = controller.controllerGetSheetNames();
            if (sheetNames.length == 0) {
                JOptionPane.showMessageDialog(null, "Нет доступных листов для экспорта.", "Ошибка", JOptionPane.ERROR_MESSAGE);
                return;
            }

            Object selectedObject = JOptionPane.showInputDialog(
                    null,
                    "Выберите лист для экспорта:",
                    "Выбор листа",
                    JOptionPane.QUESTION_MESSAGE,
                    null,
                    sheetNames,
                    sheetNames[0]
            );

            if (selectedObject == null) {
                JOptionPane.showMessageDialog(null, "Лист не выбран.", "Ошибка", JOptionPane.ERROR_MESSAGE);
                return;
            }

            int selectedSheetIndex = -1;
            for (int i = 0; i < sheetNames.length; i++) {
                if (sheetNames[i].equals(selectedObject)) {
                    selectedSheetIndex = i;
                    break;
                }
            }

            controller.controllerChooseSheet(selectedSheetIndex);

            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Сохранить файл");
            int result = fileChooser.showSaveDialog(null);

            if (result == JFileChooser.APPROVE_OPTION) {
                String destFilePath = fileChooser.getSelectedFile().getAbsolutePath() + ".xlsx";

                try {
                    controller.controllerCreateExportBook(destFilePath);
                    JOptionPane.showMessageDialog(null, "Показатели успешно экспортированы!");
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, "Ошибка при экспорте файла: " + ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        exitButton = new JButton("Выход из программы");
        add(exitButton);

        exitButton.addActionListener((ActionEvent e) -> {
            System.exit(0);
        });
        setBounds(650, 250, 600, 400);
        setVisible(true);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }

    public void setExportButtonEnabled(boolean enabled) {
        exportButton.setEnabled(enabled);
    }

}
