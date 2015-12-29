package com.rex.poi.excel.util;

public class DoubleUtil {

	public static Double getValue(String content) {
		if (content == null || "".equals(content)) {
			return 0.0;
		}
		
		try {
			Double.valueOf(content);
		} catch (NumberFormatException e) {
			return 0.0;
		}
		return Double.valueOf(content);
	}
	
	public static void main(String[] args) {
		try {
			Double.valueOf("1a1");
		} catch (NumberFormatException e) {
			System.out.println(0.0);
		}
		
		System.out.println(Double.valueOf("11"));
	}
	
}
