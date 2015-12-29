package com.rex.poi.excel.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.rex.poi.excel.model.ExcelModel;
import com.rex.poi.excel.model.MonthCode;

public class RandomUtil {
	private static SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
	
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
	}
	
	public static Map<MonthCode, List<Double>> getDataDescOrderPerMonth(List<ExcelModel> datas) {
		for (ExcelModel model : datas) {
			int month = getMonth(model.getColumnB());
			if (MonthCode.JAN.ordinal() == month) {
				janList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.FEB.ordinal() == month) {
				febList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.MAR.ordinal() == month) {
				marList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.APR.ordinal() == month) {
				aprList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.MAY.ordinal() == month) {
				mayList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.JUN.ordinal() == month) {
				junList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.JUL.ordinal() == month) {
				julList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.AUG.ordinal() == month) {
				augList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.SEP.ordinal() == month) {
				sepList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.OCT.ordinal() == month) {
				octList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.NOV.ordinal() == month) {
				novList.add(DoubleUtil.getValue(model.getColumnD()));
			} else if (MonthCode.DEC.ordinal() == month) {
				decList.add(DoubleUtil.getValue(model.getColumnD()));
			}
		}
		
		getDataListDescOrder(janList);
		getDataListDescOrder(febList);
		getDataListDescOrder(marList);
		getDataListDescOrder(aprList);
		getDataListDescOrder(mayList);
		getDataListDescOrder(junList);
		getDataListDescOrder(julList);
		getDataListDescOrder(augList);
		getDataListDescOrder(sepList);
		getDataListDescOrder(octList);
		getDataListDescOrder(novList);
		addToMapIfNotEmpty(MonthCode.JAN, janList);
		addToMapIfNotEmpty(MonthCode.FEB, febList);
		addToMapIfNotEmpty(MonthCode.MAR, marList);
		addToMapIfNotEmpty(MonthCode.APR, aprList);
		addToMapIfNotEmpty(MonthCode.MAY, mayList);
		addToMapIfNotEmpty(MonthCode.JUN, junList);
		addToMapIfNotEmpty(MonthCode.JUL, julList);
		addToMapIfNotEmpty(MonthCode.AUG, augList);
		addToMapIfNotEmpty(MonthCode.SEP, sepList);
		addToMapIfNotEmpty(MonthCode.OCT, octList);
		addToMapIfNotEmpty(MonthCode.NOV, novList);
		addToMapIfNotEmpty(MonthCode.DEC, decList);
		
		return monthMaps;
	}
	
	private static void getDataListDescOrder(List<Double> dataList) {
		Collections.sort(dataList);
		Collections.reverse(dataList);
	}

	private static void addToMapIfNotEmpty(MonthCode code, List<Double> dataList) {
		if (dataList != null && !dataList.isEmpty()) {
			monthMaps.put(code, dataList);
		}
	}

	private static int getMonth(String columnB) {
		Calendar calendar = Calendar.getInstance();
		try {
			calendar.setTime(dateFormat.parse(columnB));
		} catch (ParseException e) {
			calendar.setTime(new Date());
		}
		return calendar.get(Calendar.MONTH);
	}

	public static void main(String[] args) throws Exception {
		System.out.println(MonthCode.FEB.name().equals("FEB"));
		
	}
}
