package com.rex.poi.excel.service;

import java.io.File;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import com.rex.poi.excel.model.ExcelModel;

public interface ExcelService {
	public List<File> getFiles(String path);
	
	public List<ExcelModel> getDatas(File file);
	
	public Workbook dealWithDatas(List<ExcelModel> datas, File source, String template, boolean limited);
	
	public void writeDataToExcel(Workbook workbook, String fileName, String outputPath); 
}
