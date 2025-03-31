
package com.mycompany.lab1.controller;

import com.mycompany.lab1.model.DataCalculations;
import com.mycompany.lab1.view.GUI;
import com.mycompany.lab1.model.DataExporter;
import com.mycompany.lab1.model.DataStorage;
import java.io.IOException;

public class Controller implements IController {
    
    private static DataStorage storage; 
    private static GUI view;
    private DataExporter exporter;
    public static void start(){
        view = new GUI(new Controller());
    }
    
    @Override
    public void controllerImportData(String filePath) throws IOException {
        storage = DataStorage.importFromExcel(filePath);
        for(String s: storage.getSheetNames()){
            System.out.println(s);
        }
    }
    @Override
    public void controllerChooseSheet(int index) {
        exporter = new DataExporter();
        exporter.setSheetIndex(index);
        System.out.println("Выбран лист с индексом: " + index);
    }
    
    @Override
    public void controllerCreateExportBook(String dest) throws IOException {
        if (storage == null) {
            throw new IllegalStateException("Данные не загружены. Сначала вызовите controllerImportData.");
        }
        int sheetIndex = exporter.getSheetIndex(); 
        DataCalculations.createMap(storage, sheetIndex); 
        DataExporter.pasteCalculations(storage, dest, sheetIndex);
    }
    @Override
    public DataStorage getStorage() {
        return storage;
    }
    @Override
    public String[] controllerGetSheetNames(){
        return storage.getSheetNames();
        
    }

    
}
