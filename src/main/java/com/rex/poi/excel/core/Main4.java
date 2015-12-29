package com.rex.poi.excel.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.rex.poi.excel.model.ExcelModel;

public class Main4 {
	private static final String PATH = "E:\\Rex\\01_Project\\60_TmpCode\\99_Minyee\\input\\20151225\\";
	private static final String OUTPUT_PATH = "E:\\Rex\\01_Project\\60_TmpCode\\99_Minyee\\output\\";
	private static final Integer BEGIN_ROW = 8;
	private static final String DEFAULT_YEAR = "2014";
	
	private static final String[] FILTER_COLUMN_E = {"期末结转"};
	private static final Map<String, Integer> limitRecordNum = new HashMap<String, Integer>(); 
	private static final Integer DEFAULT_RECORD_NUM = 20222;
	private static List<Double> largeRecords;
	private static List<ExcelModel> dataList = new ArrayList<ExcelModel>();
	
	static {
//		limitRecordNum.put("应付职工薪酬", 20);
//		limitRecordNum.put("应交税金", 20);
//		limitRecordNum.put("2013", 20);
//		limitRecordNum.put("2014", 20);
	}
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		List<File> fileList = new Main4().readFiles(PATH);
		for (File file : fileList) {
			List<ExcelModel> list = new Main4().readExcelData(file);
			new Main4().writeExcelData(list, file);
		}
	}

	private List<File> readFiles(String inputPath) {
		File root = new File(inputPath);
		File[] files = root.listFiles();
		List<File> fileList = new ArrayList<File>();
		for (File file : files) {
			if (file != null && file.isFile()) {
				fileList.add(file);
			}
		}
		return fileList;
	}

	private void writeExcelData(List<ExcelModel> list, File file) throws FileNotFoundException, IOException {
		File template = new File(PATH + "template.xls");  
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(template));
        Workbook workbook = new HSSFWorkbook(fs);
        Sheet sheet = workbook.getSheetAt(0);
        Row row;
		int index = BEGIN_ROW;
		int seqNo = 1;
		CellStyle cellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("宋体");
		cellStyle.setFont(font);
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cellStyle.setWrapText(false);
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		cellStyle.setBorderTop(CellStyle.BORDER_THIN);
		
		CellStyle cellStyle2 = workbook.createCellStyle();
		font.setFontHeightInPoints((short) 11);
		cellStyle2.setFont(font);
		cellStyle2.setAlignment(CellStyle.ALIGN_GENERAL);
		cellStyle2.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
		cellStyle2.setWrapText(false);
		cellStyle2.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle2.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle2.setBorderRight(CellStyle.BORDER_THIN);
		cellStyle2.setBorderTop(CellStyle.BORDER_THIN);
		
		if (list != null && !list.isEmpty()) {
			int maxNum = getMaxNum(file.getName());
			String year = getYear(file.getName());
//			initLargeRecords(list, maxNum);
			
//			initDataList(list, maxNum);
			for (ExcelModel data : list) {
				boolean isValid = true;
				for (String filter : FILTER_COLUMN_E) {
					if (data.getColumnE().contains(filter)) {
						isValid = false;
						continue;
					}
				}
				
				if (data.getColumnA() == null || "".equals(data.getColumnA())
						|| data.getColumnB() == null || "".equals(data.getColumnB())
						|| data.getColumnE() == null || "".equals(data.getColumnE())) {
					isValid = false;
				}
				if (!isValid) {
					continue;
				}
				row = sheet.createRow(index++);
				if (row == null) {
					continue;
				}
				row.createCell(0).setCellValue(seqNo++);
				row.createCell(1).setCellValue(year + "/" + data.getColumnA() + "/" + data.getColumnB());
				row.createCell(2).setCellValue(data.getColumnC() + "-" + data.getColumnD());
				row.createCell(3).setCellValue(data.getColumnE());
				if (data.getColumnF() == null || "".equals(data.getColumnF())
						&&(data.getColumnG() == null || "".equals(data.getColumnG()))) {
					row.createCell(4).setCellValue(0.0);
				} else if (data.getColumnF() == null || "".equals(data.getColumnF())) {
					row.createCell(4).setCellValue(Double.valueOf(data.getColumnG()));
				} else {
					row.createCell(4).setCellValue(Double.valueOf(data.getColumnF()));
				}
				row.createCell(5).setCellValue("");
				row.createCell(6).setCellValue("");
				row.createCell(7).setCellValue("");
				row.createCell(8).setCellValue("");
				row.getCell(0).setCellStyle(cellStyle);
				row.getCell(1).setCellStyle(cellStyle2);
				row.getCell(2).setCellStyle(cellStyle2);
				row.getCell(3).setCellStyle(cellStyle2);
				row.getCell(4).setCellStyle(cellStyle2);
				row.getCell(5).setCellStyle(cellStyle2);
				row.getCell(6).setCellStyle(cellStyle2);
				row.getCell(7).setCellStyle(cellStyle2);
				row.getCell(8).setCellStyle(cellStyle2);
			}
		}
		
		FileOutputStream out = new FileOutputStream(OUTPUT_PATH + "记账凭证-" + file.getName());  
        workbook.write(out);  
        workbook.close();
        out.close();  
		
	}

	private void initDataList(List<ExcelModel> list, int maxNum) {
		if (list != null && !list.isEmpty()) {
			int index = BEGIN_ROW;
			boolean cheked = false;
			dataList = new ArrayList<ExcelModel>();
			Set<String> monthSet = new HashSet<String>();
			int time = 0;
			while (maxNum > dataList.size() && time < 3) {
				for (ExcelModel data : list) {
					if (index > BEGIN_ROW + maxNum - 1) {
						time = 3;
						break;
					}
					if (!checkIfValidData(data)) {
						continue;
					}
					
					if (!cheked && monthSet.contains(data.getColumnA())) {
						continue;
					}
					monthSet.add(data.getColumnA());
					if (!dataList.contains(data)) {
						index++;
						dataList.add(data);
					}
				}
				cheked = true;
				time++;
			}
		}
		
		Collections.sort(dataList, new DateComparator());
	}

	private String getYear(String name) {
		int index = name.indexOf("201");
		if (index == -1) {
			return DEFAULT_YEAR;
		}
		String year = name.substring(index, index + 4);
		return year;
	}

	private void initLargeRecords(List<ExcelModel> list, int maxNum) {
		if (list != null && !list.isEmpty()) {
			largeRecords = new ArrayList<Double>();
			List<Double> amountList = new ArrayList<Double>();
			for (ExcelModel data : list) {
				if (data.getColumnC() == null || "".equals(data.getColumnC())) {
					continue;
				}
				
				if ((data.getColumnF() == null || "".equals(data.getColumnF()))
						&& (data.getColumnG() == null || "".equals(data.getColumnG()))) {
					continue;
				}
				
				if (data.getColumnF() == null || "".equals(data.getColumnF())) {
					amountList.add(Math.abs(Double.valueOf(data.getColumnG())));
				} else {
					amountList.add(Math.abs(Double.valueOf(data.getColumnF())));
				}
			}
			
			Collections.sort(amountList);
			Collections.reverse(amountList);
			largeRecords = amountList.subList(0, Math.min(maxNum * 2, amountList.size()));
		}
	}

	private int getMaxNum(String name) {
		for (String key : limitRecordNum.keySet()) {
			if (name.contains(key)) {
				return limitRecordNum.get(key);
			}
		}
		return DEFAULT_RECORD_NUM;
	}

	private boolean checkIfValidData(ExcelModel data) {
		if (data.getColumnA() == null || "".equals(data.getColumnA())
				|| data.getColumnB() == null || "".equals(data.getColumnB())
				|| data.getColumnE() == null || "".equals(data.getColumnE())) {
			return false;
		}
		
		for (String filter : FILTER_COLUMN_E) {
			if (data.getColumnE().contains(filter)) {
				return false;
			}
		}
		
		if ((data.getColumnF() == null || "".equals(data.getColumnF()))
				&& (data.getColumnG() == null || "".equals(data.getColumnG()))) {
			return false;
		}
		
		Double amount = 0.0;
		if (data.getColumnF() == null || "".equals(data.getColumnF())) {
			amount = Math.abs(Double.valueOf(data.getColumnG()));
		} else {
			amount = Math.abs(Double.valueOf(data.getColumnF()));
		}
		if (!largeRecords.contains(amount)) {
			return false;
		}
		
		return true;
	}

	private List<ExcelModel> readExcelData(File file) throws FileNotFoundException, IOException {
		Workbook workbook = new HSSFWorkbook(new FileInputStream(file));
		Sheet sheet = workbook.getSheetAt(0);
		Row row;
		List<ExcelModel> list = new ArrayList<ExcelModel>();
		for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row != null) {
				ExcelModel model = new ExcelModel();
				model.setColumnA(getCellValue(row.getCell(0)));
				model.setColumnB(getCellValue(row.getCell(1)));
				model.setColumnC(getCellValue(row.getCell(2)));
				model.setColumnD(getCellValue(row.getCell(3)));
				model.setColumnE(getCellValue(row.getCell(4)));
				model.setColumnF(getCellValue(row.getCell(5)));
				model.setColumnG(getCellValue(row.getCell(6)));
				
				list.add(model);
			}
		}
        workbook.close();

		return list;
	}

	private String getCellValue(Cell cell) {
		if (null != cell) {     
            switch (cell.getCellType()) {     
            case Cell.CELL_TYPE_NUMERIC: // 数字     
            	return formatVal(cell.getNumericCellValue());     
            case Cell.CELL_TYPE_STRING: // 字符串     
            	return formatVal(cell.getStringCellValue());     
            case Cell.CELL_TYPE_BOOLEAN: // Boolean     
            	return formatVal(cell.getBooleanCellValue());     
            case Cell.CELL_TYPE_FORMULA: // 公式     
            	return cell.getCellFormula();
            case Cell.CELL_TYPE_BLANK: // 空值     
            	return "";
            case Cell.CELL_TYPE_ERROR: // 故障     
                break;     
            default:     
                break;     
            }     
        }
		return null;  
	}

	private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
	private DecimalFormat df = new DecimalFormat("###0.####");
	private String formatVal(Object val) {
		if (isNotEmpty(val)) {
			if(val instanceof String)
				return (String)val;
			else if(val instanceof Date)
				return sdf.format((Date)val);
			else if(val instanceof Double)
				return df.format((Double)val);
			else if(val instanceof BigDecimal)
				return df.format((BigDecimal)val);
			else if(val instanceof Integer)
				return ((Integer)val).toString();
			else if(val instanceof Boolean)
				return (((Boolean) val).booleanValue()==true?"Y":"N");
		}
		return "";
	}
	
	private boolean isNotEmpty(Object object) {
		if (object != null) {
			if (object instanceof String) {
				return (!"".equals((String) object));
			}
			return true;
		}
		return false;
	}
	
	private class DateComparator implements Comparator<ExcelModel> {

		private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-mm-dd");
		
		@Override
		public int compare(ExcelModel o1, ExcelModel o2) {
			String o1Date = "";
			String o2Date = "";
			try {
				o1Date = sdf.format(sdf.parse("1900" + "-" + o1.getColumnA() + "-" + o1.getColumnB()));
				o2Date = sdf.format(sdf.parse("1900" + "-" + o2.getColumnA() + "-" + o2.getColumnB()));
			} catch (ParseException e) {
				e.printStackTrace();
			}
			
			
			return o1Date.compareTo(o2Date);
		}
		
	}
}
