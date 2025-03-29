package com.mycompany.lab1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class DataImporter {

    //данные с 1 листа
    public static DataStorage importFromExcel(String file, int listIndex) throws IOException {
        DataStorage ds = new DataStorage();
        Workbook data = new XSSFWorkbook(new FileInputStream(file));
        Sheet sheet = data.getSheetAt(listIndex);
        FormulaEvaluator evaluator = data.getCreationHelper().createFormulaEvaluator();
        // Получаем первую строку с названиями выборок
        Row sampleNamesRow = sheet.getRow(0);
        if (sampleNamesRow == null) {
            throw new IllegalArgumentException("Первая строка пустая");
        }

        // Перебираем столбцы
        for (int k = 0; k < sampleNamesRow.getLastCellNum(); k++) {
            Cell cell = sampleNamesRow.getCell(k);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                List<Double> tempSampleStorage = new ArrayList<>();
                
                // Перебираем строки для текущего столбца
                Iterator<Row> rowIter = sheet.iterator();
                String currSampleName = rowIter.next().getCell(k).getStringCellValue();
//                rowIter.next(); // Пропускаем первую строку с названиями выборок

                while (rowIter.hasNext()) {
                    Row row = rowIter.next();
                    if (isRowEmpty(row)) continue;

                    Cell valueCell = row.getCell(k);
                    if (valueCell != null && valueCell.getCellType() == CellType.NUMERIC) {
                        double currValue = valueCell.getNumericCellValue();
                        tempSampleStorage.add(currValue);
                    } else if(valueCell.getCellType() == CellType.FORMULA){
                        CellValue evaluatedValue = evaluator.evaluate(valueCell);
                        switch (evaluatedValue.getCellType()) {
                            case NUMERIC:
                                double currValue = evaluatedValue.getNumberValue();
                                tempSampleStorage.add(currValue);
                                break;
                            default:
                                System.out.println("Неизвестный тип формулы");
                        }
                    }
                }
                
                // Добавляем данные текущей выборки в общий список
                ds.setSampleData(currSampleName, tempSampleStorage);
            }
        }

        data.close();
        return ds;
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