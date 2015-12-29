package com.rex.poi.excel.util;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FileUtil {

	public static final String DEFAULT_FILE_YEAR = "2015";
	
	public static List<File> getFiles(String path) {
		if (path == null || "".equals(path)) {
			return null;
		}

		File root = new File(path);
		return getFile(root);
	}

	private static List<File> getFile(File root) {
		List<File> fileList = new ArrayList<File>();
		if (root.isFile()) {
			fileList.add(root);
			return fileList;
		}
		
		File[] files = root.listFiles();
		for (File file : files) {
			if (file.isDirectory()) {
				fileList.addAll(getFile(file));
			}
			fileList.add(file);
		}
		return fileList;
	}

	
	public static String getFileExtension(File file) {
		String filename = file.getName();
		int dot = filename.lastIndexOf(".");
		if ((dot > -1) && (dot < (filename.length() - 1))) { 
            return filename.substring(dot + 1); 
        }
		return filename;
	}

	public static String getFileYear(String fileName) {
		if (fileName == null || "".equals(fileName)
				|| !hasDigit(fileName)) {
			return DEFAULT_FILE_YEAR;
		}
		
		String year = fileName.substring(0, 4);
		if (isDigit(year)) {
			return year;
		} else {
			int dot = fileName.lastIndexOf(".");
			year = fileName.substring(dot - 4, dot);
			if (isDigit(year)) {
				return year;
			}
		}
		
		return DEFAULT_FILE_YEAR;
	}
	
	public static boolean hasDigit(String content){ 
		boolean flag = false;
		Pattern p = Pattern.compile(".*\\d+.*");
		Matcher m = p.matcher(content);
		if (m.matches()) {
			flag = true;
		}

		return flag;

	}
	
	public static boolean isDigit(String strNum) {  
	    Pattern pattern = Pattern.compile("[0-9]{1,}");  
	    Matcher matcher = pattern.matcher((CharSequence) strNum);  
	    return matcher.matches();  
	}
	
	public static boolean beginWithDigit(String strNum) {
		Pattern pattern = Pattern.compile("^(\\d+)(.*)");
		Matcher matcher = pattern.matcher(strNum);
		if (matcher.matches()) {//Êý×Ö¿ªÍ·
			return true;
		}
		return false;
	}

}
