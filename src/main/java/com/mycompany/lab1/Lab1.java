/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.lab1;
//import static DataImporter.importFromExcel;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
/**
 *
 * @author nsoko
 */
public class Lab1 {

     public static void main(String[] args) throws IOException {
        
//        String file = "C:\\Users\\nsoko\\Downloads\\Р›Р°Р±Р°_1 РѕР±СЂР°Р·С†С‹ РґР°РЅРЅС‹С….xlsx";
//        String dest = "C:\\Users\\nsoko\\OneDrive\\Desktop\\testEmpty9.xlsx";

        String file = "C:\\Users\\Nastya\\Downloads\\Лаба_1 образцы данных.xlsx";
        String dest = "C:\\Users\\Nastya\\Desktop\\test16_colored_full.xlsx";
        
        Map<String,Map<String, Double>> calcMap = DataExporter.createMap(file, 2);
        
        DataExporter.pasteCalculations(file, 2, dest, calcMap);
//        System.out.println(calcMap);
        
        DataStorage ds = DataImporter.importFromExcel(file, 2);
        double[][] covMatrix = DataCalculations.calculateCovarianceMatrix(ds);
        for(int i = 0; i <covMatrix.length; i++){
            for(int j = 0; j < covMatrix[i].length; j++){
                    System.out.print(covMatrix[i][j]+ " ");
            }
            System.out.println("");
        }
    }
}
