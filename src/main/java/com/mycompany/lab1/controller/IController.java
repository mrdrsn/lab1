
package com.mycompany.lab1.controller;

import com.mycompany.lab1.model.DataStorage;
import java.io.IOException;

public interface IController {
    void controllerImportData(String filePath) throws IOException;
    void controllerChooseSheet(int index);
    void controllerCreateExportBook(String dest) throws IOException;
    DataStorage getStorage();
    String[] controllerGetSheetNames();
}