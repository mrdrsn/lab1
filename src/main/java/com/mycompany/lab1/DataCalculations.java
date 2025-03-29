package com.mycompany.lab1;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.math3.distribution.NormalDistribution;
import org.apache.commons.math3.stat.descriptive.*;
import org.apache.commons.math3.stat.correlation.*;

public class DataCalculations {
    
    public static Map<String,Double> calculateGeometricMean(DataStorage sheetData) {
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry : sheetData.getSheetData().entrySet()){
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            DescriptiveStatistics d = new DescriptiveStatistics();
            
            for (double value : sample) {
                d.addValue(value);
            }
            double result = d.getGeometricMean();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    public static Map<String, Double> calculateArithmeticMean(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getMean();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    public static Map<String,Double> calculateStandardDeviation(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getStandardDeviation();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    
    public static Map<String,Double> calculateSampleRange(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getMax() - d.getMin();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    
    public static Map<String,Double> calculateSize(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getN();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    
    public static Map<String,Double> calculateMin(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getMin();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    public static Map<String,Double> calculateMax(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getMax();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    public static Map<String,Double> calculateVariance(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = d.getVariance();
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
    
    
    public static Map<String,Double> calculateMarginOfError(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double mean = d.getMean(); // Среднее значение
        double stdDev = d.getStandardDeviation(); // Стандартное отклонение
        int n = (int) d.getN(); // Размер выборки

        // Критическое значение z (нормальное распределение)
        NormalDistribution normalDist = new NormalDistribution();
        double confidenceLevel = 0.95;
        double alpha = 1 - confidenceLevel; // Уровень значимости
        double z = normalDist.inverseCumulativeProbability(1 - alpha / 2);

        // Расчет доверительного интервала
        double marginOfError = z * (stdDev / Math.sqrt(n));
        double lowerBound = mean - marginOfError;
        double upperBound = mean + marginOfError;
        resultOfCalculations.put(sampleName, marginOfError);
        }
        return resultOfCalculations;
    }
    
    public static Map<String,Double> calculateVarCoeff(DataStorage sheetData){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: sheetData.getSheetData().entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double result = (d.getStandardDeviation()/d.getMean())*100;
            resultOfCalculations.put(sampleName, result);
        }
        return resultOfCalculations;
    }
     public static double[][] calculateCovarianceMatrix(DataStorage sheetData) {
            
            // Определяем количество выборок и их размер
            int numSamples = sheetData.getSheetData().size();
            int sampleSize = -1;

            // Создаем двумерный массив для хранения данных
            double[][] dataArray = new double[numSamples][];

            // Заполняем массив данными из Map
            int index = 0;
            for (List<Double> sample : sheetData.getSheetData().values()) {
                if (sampleSize == -1) {
                    sampleSize = sample.size();
                } else if (sample.size() != sampleSize) {
                    throw new IllegalArgumentException("Все выборки должны иметь одинаковый размер");
                }

                // Преобразуем List<Double> в массив double[]
                dataArray[index] = sample.stream().mapToDouble(Double::doubleValue).toArray();
                index++;
            }

            // Транспонируем массив, чтобы столбцы соответствовали выборкам
            double[][] transposedData = transpose(dataArray);

            // Создаем объект Covariance и рассчитываем ковариационную матрицу
            Covariance covariance = new Covariance(transposedData);
            return covariance.getCovarianceMatrix().getData();
    }

    private static double[][] transpose(double[][] matrix) {
        int rows = matrix.length;
        int cols = matrix[0].length;
        double[][] transposed = new double[cols][rows];

        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < cols; j++) {
                transposed[j][i] = matrix[i][j];
            }
        }

        return transposed;
    }
    
}
