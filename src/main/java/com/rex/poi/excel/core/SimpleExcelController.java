package com.rex.poi.excel.core;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Workbook;

import com.rex.poi.excel.model.ExcelModel;
import com.rex.poi.excel.service.ExcelService;
import com.rex.poi.excel.service.SimpleExcelServiceImpl;

public class SimpleExcelController {
	private static String inputPath = "";
	private static String outputPath = "";
	private static String simpleTemplatePath = "";
	
	public static void main(String[] args) {
		init();
		ExcelService excelService = new SimpleExcelServiceImpl();
		List<File> fileList = excelService.getFiles(inputPath);
		for (File file : fileList) {
			List<ExcelModel> datas = excelService.getDatas(file);
			Workbook workbook = excelService.dealWithDatas(datas, file, simpleTemplatePath, false);
			excelService.writeDataToExcel(workbook, file.getName(), outputPath);
		}
	}

	private static void init() {
		Properties prop =  new Properties();    
		InputStream in = Object.class.getResourceAsStream("/system.properties");    
		try  {    
			prop.load(in);    
			inputPath = prop.getProperty( "inputPath" ).trim();    
			outputPath = prop.getProperty( "outputPath" ).trim();    
			simpleTemplatePath = prop.getProperty( "simpleTemplatePath" ).trim();    
		}  catch  (IOException e) {    
			e.printStackTrace();    
		}
	}
}
