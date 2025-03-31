package com.mycompany.lab1.model;

import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.math3.distribution.NormalDistribution;
import org.apache.commons.math3.stat.descriptive.*;
import org.apache.commons.math3.stat.correlation.*;

public class DataCalculations {
    static final String[] statisticNeeds = {"Среднее геометрическое", "Среднее арифметическое", "Оценка стандратного отклонения", "Размах",
        "Количество элементов", "Коэффициент вариации", "Доверительный интервал мат. ожидания", "Оценка дисперсии", "Максимум", "Минимум"};
    
    static Map<String, Map<String, Double>> statisticsMap  = new LinkedHashMap<>();
    
    
    public static Map<String,Double> calculateGeometricMean(DataStorage fileData, int sheetIndex) {
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry : fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String, Double> calculateArithmeticMean(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateStandardDeviation(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateSampleRange(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateSize(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateMin(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateMax(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateVariance(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static Map<String,Double> calculateMarginOfError(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
            DescriptiveStatistics d = new DescriptiveStatistics();
            List<Double> sample = (List<Double>) entry.getValue();
            String sampleName = (String) entry.getKey();
            for(double value: sample){
                d.addValue(value);
            }
            double mean = d.getMean(); 
        double stdDev = d.getStandardDeviation();
        int n = (int) d.getN();

        NormalDistribution normalDist = new NormalDistribution();
        double confidenceLevel = 0.95;
        double alpha = 1 - confidenceLevel;
        double z = normalDist.inverseCumulativeProbability(1 - alpha / 2);

        double marginOfError = z * (stdDev / Math.sqrt(n));
        resultOfCalculations.put(sampleName, marginOfError);
        }
        return resultOfCalculations;
    }
    
    public static Map<String,Double> calculateVarCoeff(DataStorage fileData, int sheetIndex){
        Map<String,Double> resultOfCalculations = new HashMap<>();
        for(Map.Entry entry: fileData.getSheetData(sheetIndex).entrySet()){
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
    
    public static double[][] calculateCovarianceMatrix(DataStorage fileData, int sheetIndex) {
            
            int numSamples = fileData.getSheetData(sheetIndex).size();
            int sampleSize = -1;

            double[][] dataArray = new double[numSamples][];

            int index = 0;
            for (List<Double> sample : fileData.getSheetData(sheetIndex).values()) {
                if (sampleSize == -1) {
                    sampleSize = sample.size();
                } else if (sample.size() != sampleSize) {
                    throw new IllegalArgumentException("Все выборки должны иметь одинаковый размер");
                }

                dataArray[index] = sample.stream().mapToDouble(Double::doubleValue).toArray();
                index++;
            }

            double[][] transposedData = transpose(dataArray);

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
    
    public static void createMap(DataStorage storage, int sheetIndex) throws IOException {

        Map<String, Double> geometricMeanMap = calculateGeometricMean(storage, sheetIndex);
        Map<String, Double> arithmeticMeanMap = calculateArithmeticMean(storage, sheetIndex);
        Map<String, Double> standardDeviationMap = calculateStandardDeviation(storage, sheetIndex);
        Map<String, Double> sampleRangeMap = calculateSampleRange(storage, sheetIndex);
        Map<String, Double> sizeMap = calculateSize(storage, sheetIndex);
        Map<String, Double> minMap = calculateMin(storage, sheetIndex);
        Map<String, Double> maxMap = calculateMax(storage, sheetIndex);
        Map<String, Double> marginOfErrorMap = calculateMarginOfError(storage, sheetIndex);
        Map<String, Double> varianceMap = calculateVariance(storage, sheetIndex);
        Map<String, Double> varCoeffMap = calculateVarCoeff(storage, sheetIndex);

        statisticsMap.put(statisticNeeds[0], geometricMeanMap); // Среднее геометрическое
        statisticsMap.put(statisticNeeds[1], arithmeticMeanMap); // Среднее арифметическое
        statisticsMap.put(statisticNeeds[2], standardDeviationMap); // Оценка стандартного отклонения
        statisticsMap.put(statisticNeeds[3], sampleRangeMap); // Размах
        statisticsMap.put(statisticNeeds[4], sizeMap); // Количество элементов
        statisticsMap.put(statisticNeeds[5], varCoeffMap); // Коэффициент вариации
        statisticsMap.put(statisticNeeds[6], marginOfErrorMap); //значение сдвига для доверительного интервала
        statisticsMap.put(statisticNeeds[7], varianceMap); // Оценка дисперсии
        statisticsMap.put(statisticNeeds[8], maxMap); // Максимум
        statisticsMap.put(statisticNeeds[9], minMap); // Минимум
    }   
}
