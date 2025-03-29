package com.mycompany.lab1;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
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
    public static void calculateCovariation(DataStorage sheetData){
        Covariance cov = new Covariance();
        
        
    }
    
}
