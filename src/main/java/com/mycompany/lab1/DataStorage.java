package com.mycompany.lab1;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

//хранилище для информации с 1 листа!

public class DataStorage {
    private Map<String, List<Double>> sheetData = new HashMap<>();

    public Map<String, List<Double>> getSheetData(){
        return this.sheetData;
    }
    public List<Double> getSampleData(String key){
        List<Double> sample = new ArrayList<>();
        return getSheetData().get(key);
    }
    
    public String[] getSampleNames(){
        String[] array = getSheetData().keySet().toArray(new String[0]);
        return array;
    }
    
    public void setSampleData(String key, List<Double> values){
        this.sheetData.put(key, values);
    }
}