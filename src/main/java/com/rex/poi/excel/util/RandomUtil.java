package com.rex.poi.excel.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.rex.poi.excel.model.BaseModel;
import com.rex.poi.excel.model.MonthCode;

public class RandomUtil {
	private static List<Double> janList;
	private static List<Double> febList;
	private static List<Double> marList;
	private static List<Double> aprList;
	private static List<Double> mayList;
	private static List<Double> junList;
	private static List<Double> julList;
	private static List<Double> augList;
	private static List<Double> sepList;
	private static List<Double> octList;
	private static List<Double> novList;
	private static List<Double> decList;
	
	private static Map<MonthCode, List<Double>> monthMaps;
	
	public static void initList() {
		janList = new ArrayList<Double>();
		febList = new ArrayList<Double>();
		marList = new ArrayList<Double>();
		aprList = new ArrayList<Double>();
		mayList = new ArrayList<Double>();
		junList = new ArrayList<Double>();
		julList = new ArrayList<Double>();
		augList = new ArrayList<Double>();
		sepList = new ArrayList<Double>();
		octList = new ArrayList<Double>();
		novList = new ArrayList<Double>();
		decList = new ArrayList<Double>();
	}
	
	public static void initMonthMap() {
		monthMaps = new HashMap<MonthCode, List<Double>>();
		monthMaps.put(MonthCode.JAN, janList);
		monthMaps.put(MonthCode.FEB, febList);
		monthMaps.put(MonthCode.MAR, marList);
		monthMaps.put(MonthCode.APR, aprList);
		monthMaps.put(MonthCode.MAY, mayList);
		monthMaps.put(MonthCode.JUN, junList);
		monthMaps.put(MonthCode.JUL, julList);
		monthMaps.put(MonthCode.AUG, augList);
		monthMaps.put(MonthCode.SEP, sepList);
		monthMaps.put(MonthCode.OCT, octList);
		monthMaps.put(MonthCode.NOV, novList);
		monthMaps.put(MonthCode.DEC, decList);
	}
	
	public static void getRandomLargeData(List<BaseModel> modelList) {
		for (BaseModel model : modelList) {
			
		}
	}
	
	public static void main(String[] args) {
		System.out.println(MonthCode.FEB.name().equals("FEB"));
	}
}
