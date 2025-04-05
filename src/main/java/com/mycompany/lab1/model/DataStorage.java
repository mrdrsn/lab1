package com.mycompany.lab1.model;

import exceptions.InvalidFileFormatException;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//хранилище для информации со всех листов!

public class DataStorage {    
    
    private final Map<Integer, Map<String,List<Double>>> fileData = new HashMap<>();

    public Map<String, List<Double>> getSheetData(int sheetIndex) {
        Map<String, List<Double>> sheetData = this.fileData.get(sheetIndex);
        if (sheetData == null) {
            return new HashMap<>();
        }
        return sheetData;
    } 
    
    public String[] getSampleNames(int sheetIndex){
        String[] array = getSheetData(sheetIndex).keySet().toArray(new String[0]);
        return array;
    }
    public Map<Integer,Map<String,List<Double>>> getFileData(){
        return this.fileData;
    }
    private void setFileData(int key, Map<String,List<Double>> values){
        this.fileData.put(key, values);
    }
    
    public String[] getSheetNames() {
        return fileData.keySet().stream().map(String::valueOf).toArray(String[]::new);
    }
    
    //выгрузка всего эксель документа в единый data storage
    public static DataStorage importFromExcel(String file) throws IOException, InvalidFileFormatException {
        DataStorage ds = new DataStorage();
        try{
        Workbook data = new XSSFWorkbook(new FileInputStream(file));
        Iterator <Sheet> sheetIter = data.iterator();
        int sheetIndex = 0;
        while(sheetIter.hasNext()){
            
            Sheet sheet = sheetIter.next();
            FormulaEvaluator evaluator = data.getCreationHelper().createFormulaEvaluator();

            Row sampleNamesRow = sheet.getRow(0);
            if (sampleNamesRow == null) {
                throw new IllegalArgumentException("Первая строка пустая");
            }
            Map<String,List<Double>> sampleData = new HashMap<>();
           
            for (int k = 0; k < sampleNamesRow.getLastCellNum(); k++) {
                Cell cell = sampleNamesRow.getCell(k);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    List<Double> tempSampleStorage = new ArrayList<>();

                    Iterator<Row> rowIter = sheet.iterator();
                    String currSampleName = cell.getStringCellValue();

                    while (rowIter.hasNext()) {
                        Row row = rowIter.next();

                        if (isRowEmpty(row)) continue; //проверка на дырки в данных

                        Cell valueCell = row.getCell(k);
                        if (valueCell != null && valueCell.getCellType() == CellType.NUMERIC) {
                            double currValue = valueCell.getNumericCellValue();
                            tempSampleStorage.add(currValue);
                        } else if(valueCell.getCellType() == CellType.FORMULA){
                            CellValue evaluatedValue = evaluator.evaluate(valueCell);
                            double currValue = evaluatedValue.getNumberValue();
                            tempSampleStorage.add(currValue);
                        }
                    }
                    sampleData.put(currSampleName, tempSampleStorage);
                }
            }
            ds.setFileData(sheetIndex, sampleData);
            sheetIndex++;
        }
        data.close();
        return ds;
        } catch(NotOfficeXmlFileException ex){
            throw new InvalidFileFormatException("Файл не является допустимым файлом Excel (.xlsx)", ex);
        }
    }
    
    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }
}