package com.mycompany.lab1.model;

import com.mycompany.lab1.model.DataStorage;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class DataExporter {


    public static DataStorage storage = new DataStorage();
    private int sheetIndex;

    public void setSheetIndex(int index) {
        this.sheetIndex = index;
    }

    public int getSheetIndex() {
        return this.sheetIndex;
    }

    public static Workbook createEmptyExportBook(DataStorage storage, int sheetIndex) throws FileNotFoundException, IOException{
        
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Расчеты для выборок");

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);  
        
        XSSFColor customOrange = new XSSFColor(new java.awt.Color(255, 204, 153), null);
        ((XSSFCellStyle) cellStyle).setFillForegroundColor(customOrange);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font boldFont = wb.createFont();
        boldFont.setBold(true); 
        cellStyle.setFont(boldFont); 
        
        Row row = sheet.createRow(0);
        String[] sampleNames = storage.getSampleNames(sheetIndex);
        //создаю шапку таблицы (обозначения выборок)
        for(int i = 0; i < sampleNames.length; i++){
            Cell sampleName = row.createCell(i+1);
            sampleName.setCellValue(sampleNames[i]);
            sampleName.setCellStyle(cellStyle);
        }

        //создаю столбец с обозначением стат показателя
        for(int i = 0; i < DataCalculations.statisticNeeds.length; i++){
            Row r = sheet.createRow(i+1);
            Cell statName = r.createCell(0);
            statName.setCellValue(DataCalculations.statisticNeeds[i]);
            statName.setCellStyle(cellStyle);
        }

        Row covRow = sheet.getRow(0);
        Cell covTitle = covRow.createCell(2 + sampleNames.length); 
        covTitle.setCellValue("Матрица ковариации");

        for (int i = 0; i < sampleNames.length; i++) {
            Row rowI = sheet.getRow(0);
            Cell sampleName = rowI.createCell(i + 1 + covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка "+ sampleNames[i]);
            sampleName.setCellStyle(cellStyle);
        }

        for (int j = 0; j < sampleNames.length; j++) {
            Row rowJ = sheet.getRow(j+1);
            Cell sampleName = rowJ.createCell(covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка " + sampleNames[j]);
            sampleName.setCellStyle(cellStyle);
        }
        for (int i = 0; i <= 2+2*sampleNames.length; i++) {
            sheet.autoSizeColumn(i);
            if(i >= 1 && i <= 5){
                sheet.setColumnWidth(i, 256*20);
            }
        }
        return wb;
    }
    
    public static void pasteCalculations(DataStorage storage,String dest, int sheetIndex) throws IOException {
        Workbook wb = createEmptyExportBook(storage,sheetIndex);
        Sheet workSheet = wb.getSheet("Расчеты для выборок");
        DataCalculations.createMap(storage, sheetIndex);
        
        CellStyle centeredCellStyle = wb.createCellStyle();
        centeredCellStyle.setAlignment(HorizontalAlignment.CENTER);
        
        Iterator<Row> rowIter = workSheet.iterator();
        if (rowIter.hasNext()) {
            rowIter.next();
        }

      
        for (Map.Entry<String, Map<String, Double>> entry : DataCalculations.statisticsMap.entrySet()) {
            String statisticKey = entry.getKey(); 
            Map<String, Double> mapValue = entry.getValue();

            
            while (rowIter.hasNext()) {
                Row currRow = rowIter.next();
                Cell firstCell = currRow.getCell(0); 

                if (firstCell != null && statisticKey.equals(firstCell.getStringCellValue())) {
            
                    for (Cell headerCell : workSheet.getRow(0)) {
                        String sampleKey = headerCell.getStringCellValue(); 
                        if (mapValue.containsKey(sampleKey)) {
                            int j = headerCell.getColumnIndex(); 
                            Cell valueCell = currRow.createCell(j); 
                            if (statisticKey.equals(DataCalculations.statisticNeeds[6])) {
                                double lowerBound = DataCalculations.statisticsMap.get(DataCalculations.statisticNeeds[1]).get(sampleKey) - mapValue.get(sampleKey);
                                double upperBound = DataCalculations.statisticsMap.get(DataCalculations.statisticNeeds[1]).get(sampleKey) + mapValue.get(sampleKey);

                                String formattedLowerBound = String.format("%.4f", lowerBound);
                                String formattedUpperBound = String.format("%.4f", upperBound);

                                valueCell.setCellValue("[" + formattedLowerBound + "; " + formattedUpperBound + "]");
                            } else {
                                valueCell.setCellValue(mapValue.get(sampleKey));
                            }
                            valueCell.setCellStyle(centeredCellStyle);
                        }
                    }
                    break;
                }
            }
        }

        double[][] covMatrix = DataCalculations.calculateCovarianceMatrix(storage,sheetIndex);
        for (int i = 0; i < covMatrix.length; i++) {
            Row row = workSheet.getRow(i + 1);
            for (int j = 0; j < covMatrix[i].length; j++) {
                Cell cell = row.createCell(3 + storage.getSampleNames(sheetIndex).length + j);
                cell.setCellValue(covMatrix[i][j]);
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(dest)) {
            wb.write(fileOut);
        } finally {
            wb.close();
        }
    }
    

}
