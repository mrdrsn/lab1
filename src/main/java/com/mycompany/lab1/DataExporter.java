package com.mycompany.lab1;

import com.mycompany.lab1.DataCalculations;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class DataExporter {

    public static String[] statisticNeeds = {"Среднее геометрическое", "Среднее арифметическое", "Оценка стандратного отклонения",
                                            "Размах",
                                            "Количество элементов","Коэффициент вариации", "Доверительный этап мат. ожидания",
                                            "Оценка дисперсии",
                                            "Максимум", "Минимум"};
//    public String[] sampleName;
    public static DataStorage storage = new DataStorage();
    public static int listIndex;

    public static Workbook createEmptyExportBook(String source, int listIndex, String dest) throws FileNotFoundException, IOException{

        DataExporter.storage = DataImporter.importFromExcel(source, listIndex);
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Расчеты для выборок");

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);  
        
        XSSFColor customOrange = new XSSFColor(new java.awt.Color(255, 204, 153), null);
        ((XSSFCellStyle) cellStyle).setFillForegroundColor(customOrange);
//        cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Создаем шрифт для жирного текста
        Font boldFont = wb.createFont();
        boldFont.setBold(true); // Устанавливаем жирный шрифт
        cellStyle.setFont(boldFont); // Применяем шрифт к стилю
        
        Row row = sheet.createRow(0);
        Cell emptyCell = row.createCell(0);

        //создаю шапку таблицы (обозначения выборок)
        for(int i = 0; i < storage.getSampleNames().length; i++){
            Cell sampleName = row.createCell(i+1);
            sampleName.setCellValue(storage.getSampleNames()[i]);
            sampleName.setCellStyle(cellStyle);
        }

        //создаю столбец с обозначением стат показателя
        for(int i = 0; i < statisticNeeds.length; i++){
            Row r = sheet.createRow(i+1);
            Cell statName = r.createCell(0);
            statName.setCellValue(statisticNeeds[i]);
            statName.setCellStyle(cellStyle);
        }

        //создаю матрицу ковариации(?)
        Row covRow = sheet.getRow(0);
        Cell covTitle = covRow.createCell(2 + storage.getSampleNames().length); //5 должно вычисляться по формуле
        covTitle.setCellValue("Матрица ковариации");

        for (int i = 0; i < storage.getSampleNames().length; i++) {
            Row rowI = sheet.getRow(0); // Используем существующие строки
            Cell sampleName = rowI.createCell(i + 1 + covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка "+storage.getSampleNames()[i]);
            sampleName.setCellStyle(cellStyle);
        }

        for (int j = 0; j < storage.getSampleNames().length; j++) {
            Row rowJ = sheet.getRow(j+1); // Используем существующие строки
            Cell sampleName = rowJ.createCell(covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка " + storage.getSampleNames()[j]);
            sampleName.setCellStyle(cellStyle);
        }
        // Автоматическое изменение размера столбцов
        for (int i = 0; i <= 2+2*storage.getSampleNames().length; i++) {
            sheet.autoSizeColumn(i);
            if(i >= 1 && i <= 5){
                sheet.setColumnWidth(i, 256*20);
            }
        }
        wb.write(new FileOutputStream(dest));
        return wb;
    }

    
     public static void pasteCalculations(String source, int listIndex, String file, Map<String, Map<String, Double>> map) throws IOException {
        // Создаем пустую книгу и заполняем её структурой
        Workbook wb = createEmptyExportBook(source, listIndex, file);
        Sheet workSheet = wb.getSheet("Расчеты для выборок");
        Row rowZero = workSheet.getRow(0); // Первая строка с названиями выборок
        CellStyle centeredStyle = wb.createCellStyle();
        centeredStyle.setAlignment(HorizontalAlignment.CENTER);
        // Пропускаем первую строку (заголовки)
        Iterator<Row> rowIter = workSheet.iterator();
        if (rowIter.hasNext()) {
            rowIter.next(); // Пропускаем заголовки
        }

        // Перебираем записи в map
        for (Map.Entry<String, Map<String, Double>> entry : map.entrySet()) {
            String statisticKey = entry.getKey(); // Ключ статистики (например, "Среднее геометрическое")
            Map<String, Double> mapValue = entry.getValue(); // Значения для выборок
            
            // Перебираем строки листа
            while (rowIter.hasNext()) {
                Row currRow = rowIter.next();
                Cell firstCell = currRow.getCell(0); // Первая ячейка текущей строки

                // Проверяем, совпадает ли ключ статистики с первой ячейкой строки
                if (firstCell != null && statisticKey.equals(firstCell.getStringCellValue())) {
                    // Перебираем ячейки первой строки (выборки)
                    for (Cell headerCell : rowZero) {
                        String sampleKey = headerCell.getStringCellValue(); // Название выборки
                        if (mapValue.containsKey(sampleKey)) {
                            int j = headerCell.getColumnIndex(); // Индекс столбца
                            Cell valueCell = currRow.createCell(j); // Создаем ячейку в текущей строке
                            if(statisticKey.equals(statisticNeeds[6])){
                                
                                double lowerBound = map.get(statisticNeeds[1]).get(sampleKey) - mapValue.get(sampleKey);
                                double upperBound = map.get(statisticNeeds[1]).get(sampleKey) + mapValue.get(sampleKey);
                                
                                String formattedLowerBound = String.format("%.4f", lowerBound);
                                String formattedUpperBound = String.format("%.4f", upperBound);

                                valueCell.setCellValue("[" + formattedLowerBound + "; " + formattedUpperBound + "]");
                            }else{
                                valueCell.setCellValue(mapValue.get(sampleKey)); // Записываем значение
                            }
                            valueCell.setCellStyle(centeredStyle);
                        }
                    }
                    break; // Переходим к следующей записи в map
                }
            }
        }

        double[][] covMatrix = createCovmatrix(source, listIndex);
        for(int i = 0; i <covMatrix.length; i++){
            Row row = workSheet.getRow(i+1);
            for(int j = 0; j < covMatrix[i].length; j++){
                    Cell cell = row.createCell(3 + storage.getSampleNames().length + j);
                    cell.setCellValue(covMatrix[i][j]);
            }
        }
        // Записываем данные в файл и закрываем Workbook
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            wb.write(fileOut);
        } finally {
            wb.close();
        }
    }
 
    
    
    
    
    
    
    
    public static double[][] createCovmatrix(String source, int listIndex) throws IOException{
        DataExporter.storage = DataImporter.importFromExcel(source, listIndex);
        return DataCalculations.calculateCovarianceMatrix(storage);
    } 
     
    public static Map<String, Map<String, Double>> createMap(String source, int listIndex) throws IOException {
    // Используем LinkedHashMap для сохранения порядка добавления
    Map<String, Map<String, Double>> statisticsMap = new LinkedHashMap<>();

    DataExporter.storage = DataImporter.importFromExcel(source, listIndex);
    Map<String, Double> geometricMeanMap = DataCalculations.calculateGeometricMean(storage);
    Map<String, Double> arithmeticMeanMap = DataCalculations.calculateArithmeticMean(storage);
    Map<String, Double> standardDeviationMap = DataCalculations.calculateStandardDeviation(storage);
    Map<String, Double> sampleRangeMap = DataCalculations.calculateSampleRange(storage);
    Map<String, Double> sizeMap = DataCalculations.calculateSize(storage);
    Map<String, Double> minMap = DataCalculations.calculateMin(storage);
    Map<String, Double> maxMap = DataCalculations.calculateMax(storage);
    Map<String,Double> marginOfErrorMap = DataCalculations.calculateMarginOfError(storage);
    Map<String, Double> varianceMap = DataCalculations.calculateVariance(storage);
    Map<String, Double> varCoeffMap = DataCalculations.calculateVarCoeff(storage);

    statisticsMap.put(statisticNeeds[0], geometricMeanMap); // Среднее геометрическое
    statisticsMap.put(statisticNeeds[1], arithmeticMeanMap); // Среднее арифметическое
    statisticsMap.put(statisticNeeds[2], standardDeviationMap); // Оценка стандартного отклонения
    statisticsMap.put(statisticNeeds[3], sampleRangeMap); // Размах
    statisticsMap.put(statisticNeeds[4], sizeMap); // Количество элементов
    statisticsMap.put(statisticNeeds[5], varCoeffMap); // Коэффициент вариации
    statisticsMap.put(statisticNeeds[6], marginOfErrorMap);
    statisticsMap.put(statisticNeeds[7], varianceMap); // Оценка дисперсии
    statisticsMap.put(statisticNeeds[8], maxMap); // Максимум
    statisticsMap.put(statisticNeeds[9], minMap); // Минимум

    return statisticsMap;
}
}
