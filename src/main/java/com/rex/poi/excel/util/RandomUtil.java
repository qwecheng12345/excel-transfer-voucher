package com.rex.poi.excel.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
	
	private static final String JAN = "Jan";
	private static final String FEB = "Feb";
	private static final String MAR = "Mar";
	private static final String APR = "Apr";
	private static final String MAY = "MAY";
	private static final String JUN = "Jun";
	private static final String JUL = "Jul";
	private static final String AUG = "Aug";
	private static final String SEP = "Sep";
	private static final String OCT = "Oct";
	private static final String NOV = "Nov";
	private static final String DEC = "Dec";
	
	
	private static Map<String, List<Double>> monthMaps;
	
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
		monthMaps = new HashMap<String, List<Double>>();
		monthMaps.put(JAN, janList);
		monthMaps.put(FEB, febList);
		monthMaps.put(MAR, marList);
		monthMaps.put(APR, aprList);
		monthMaps.put(MAY, mayList);
		monthMaps.put(JUN, junList);
		monthMaps.put(JUL, julList);
		monthMaps.put(AUG, augList);
		monthMaps.put(SEP, sepList);
		monthMaps.put(OCT, octList);
		monthMaps.put(NOV, novList);
		monthMaps.put(DEC, decList);
	}
}
