package com.mycompany.lab1;

import com.mycompany.lab1.DataCalculations;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

        Row row = sheet.createRow(0);
        Cell emptyCell = row.createCell(0);

        //создаю шапку таблицы (обозначения выборок)
        for(int i = 0; i < storage.getSampleNames().length; i++){
            Cell sampleName = row.createCell(i+1);
            sampleName.setCellValue(storage.getSampleNames()[i]);
        }

        //создаю столбец с обозначением стат показателя
        for(int i = 0; i < statisticNeeds.length; i++){
            Row r = sheet.createRow(i+1);
            Cell statName = r.createCell(0);
            statName.setCellValue(statisticNeeds[i]);
        }

        //создаю матрицу ковариации(?)
        Row covRow = sheet.getRow(0);
        Cell covTitle = covRow.createCell(2 + storage.getSampleNames().length); //5 должно вычисляться по формуле
        covTitle.setCellValue("Матрица ковариации");

        for (int i = 0; i < storage.getSampleNames().length; i++) {
            Row rowI = sheet.getRow(0); // Используем существующие строки
            Cell sampleName = rowI.createCell(i + 1 + covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка " + storage.getSampleNames()[i]);
        }

        for (int j = 0; j < storage.getSampleNames().length; j++) {
            Row rowJ = sheet.getRow(j+1); // Используем существующие строки
            Cell sampleName = rowJ.createCell(covTitle.getColumnIndex());
            sampleName.setCellValue("Выборка " + storage.getSampleNames()[j]);
        }
        // Автоматическое изменение размера столбцов
        for (int i = 0; i <= 2+2*storage.getSampleNames().length; i++) {
            sheet.autoSizeColumn(i);
            if(i >= 1 && i <= 5){
                sheet.setColumnWidth(i, 256*10);
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
                            valueCell.setCellValue(mapValue.get(sampleKey)); // Записываем значение
                        }
                    }
                    break; // Переходим к следующей записи в map
                }
            }
        }

        // Записываем данные в файл и закрываем Workbook
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            wb.write(fileOut);
        } finally {
            wb.close();
        }
    }
     
     
    public static Map<String,Map<String,Double>> createMap(String source, int listIndex) throws IOException{
        Map<String,Map<String,Double>> statisticsMap = new HashMap<>();
        DataExporter.storage = DataImporter.importFromExcel(source, listIndex);
        Map<String, Double> geometricMeanMap = DataCalculations.calculateGeometricMean(storage);
        Map<String, Double> arithmeticMeanMap = DataCalculations.calculateArithmeticMean(storage);
        Map<String, Double> standardDeviationMap = DataCalculations.calculateStandardDeviation(storage);
        Map<String, Double> sampleRangeMap = DataCalculations.calculateSampleRange(storage);
        Map<String, Double> sizeMap = DataCalculations.calculateSize(storage);
        Map<String, Double> minMap = DataCalculations.calculateMin(storage);
        Map<String, Double> maxMap = DataCalculations.calculateMax(storage);
        Map<String, Double> varianceMap = DataCalculations.calculateVariance(storage);
        Map<String, Double> varCoeffMap = DataCalculations.calculateVarCoeff(storage);

        statisticsMap.put(statisticNeeds[0], geometricMeanMap); // Среднее геометрическое
        statisticsMap.put(statisticNeeds[1], arithmeticMeanMap); // Среднее арифметическое
        statisticsMap.put(statisticNeeds[2], standardDeviationMap); // Оценка стандартного отклонения
        statisticsMap.put(statisticNeeds[3], sampleRangeMap); // Размах
        statisticsMap.put(statisticNeeds[4], sizeMap); // Количество элементов
        statisticsMap.put(statisticNeeds[5], varCoeffMap); // Коэффициент вариации
        statisticsMap.put(statisticNeeds[7], varianceMap); // Оценка дисперсии
        statisticsMap.put(statisticNeeds[8], maxMap); // Максимум
        statisticsMap.put(statisticNeeds[9], minMap); // Минимум

        return statisticsMap;
    }
}
