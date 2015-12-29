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

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.rex.poi.excel.model.ExcelModel;

public class Main2 {
	private static final String PATH = "E:\\Rex\\01_Project\\60_TmpCode\\99_Minyee\\input\\20151225\\";
	private static final String OUTPUT_PATH = "E:\\Rex\\01_Project\\60_TmpCode\\99_Minyee\\output\\";
	private static final Integer BEGIN_ROW = 8;
	
	private static final String[] FILTER_COLUMN_SUMMARY = {"管理费用/劳保费", "管理费用/试验费", "管理费用/质量成本"};
	private static final Map<String, Integer> limitRecordNum = new HashMap<String, Integer>(); 
	private static final Map<String, Integer> limitRecordNumBySummary = new HashMap<String, Integer>(); 
	private static final Integer DEFAULT_RECORD_NUM_BY_SUMMARY = 80;
	private static final Integer DEFAULT_RECORD_NUM = 5;
	private static final Integer DEFAULT_RECORD_NUM2 = 10;
	private static final int DEFAULT_FACTOR = 2;
	private static final String EXTENSION = ".xls";
	private static List<Double> largeRecords;
	private static List<ExcelModel> dataList;
	
	static {
		limitRecordNumBySummary.put("管理费用/劳保费", 0);
		limitRecordNumBySummary.put("管理费用/试验费", 0);
		limitRecordNumBySummary.put("管理费用/质量成本", 0);
		limitRecordNumBySummary.put("管理费用/研发费", DEFAULT_RECORD_NUM2);
		limitRecordNumBySummary.put("管理费用/业务招待费", DEFAULT_RECORD_NUM2);
		limitRecordNumBySummary.put("管理费用/职工薪酬", DEFAULT_RECORD_NUM2);
		limitRecordNumBySummary.put("管理费用/其他费用", DEFAULT_RECORD_NUM2);
		limitRecordNumBySummary.put("管理费用/聘请中介机构费", DEFAULT_RECORD_NUM2);
		limitRecordNumBySummary.put("2013", 50);
		
	}
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		List<File> fileList = new Main2().readFiles(PATH);
		for (File file : fileList) {
			String type = "管理费用";
			String year = "2013";
			List<ExcelModel> list = new ArrayList<ExcelModel>();
			
			type = "管理费用";
			year = "2013";
			list = new Main2().readExcelData(file, type, year);
			new Main2().writeExcelData(list, file, type, year);
			
			type = "管理费用";
			year = "2014";
			list.clear();
			list = new Main2().readExcelData(file, type, year);
			new Main2().writeExcelData(list, file, type, year);
			
			type = "管理费用";
			year = "2015";
			list.clear();
			list = new Main2().readExcelData(file, type, year);
			new Main2().writeExcelData(list, file, type, year);
			
			type = "销售费用";
			year = "2013";
			list.clear();
			list = new Main2().readExcelData(file, type, year);
			new Main2().writeExcelData(list, file, type, year);
			
			type = "销售费用";
			year = "2014";
			list.clear();
			list = new Main2().readExcelData(file, type, year);
			new Main2().writeExcelData(list, file, type, year);
			
			type = "销售费用";
			year = "2015";
			list.clear();
			list = new Main2().readExcelData(file, type, year);
			new Main2().writeExcelData(list, file, type, year);
		}
	}

	private List<File> readFiles(String inputPath) {
		File root = new File(inputPath);
		File[] files = root.listFiles();
		List<File> fileList = new ArrayList<File>();
		for (File file : files) {
			if (file != null && file.isFile()) {
				if (file.getName().contains("template")) {
					continue;
				}
				fileList.add(file);
			}
		}
		return fileList;
	}

	private void writeExcelData(List<ExcelModel> list, File file, String type, String year) throws FileNotFoundException, IOException {
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
			int maxNum = getMaxNum(year);
			initLargeRecords(list, maxNum);
			
			initDataList(list, maxNum);
			for (ExcelModel data : dataList) {
				row = sheet.createRow(index++);
				if (row == null) {
					continue;
				}
				row.createCell(0).setCellValue(seqNo++);
				row.createCell(1).setCellValue(data.getColumnB());
				row.createCell(2).setCellValue(data.getColumnC());
				row.createCell(3).setCellValue(data.getColumnD());
				row.createCell(4).setCellValue(Double.valueOf(data.getColumnE()));
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
		
		String filename = file.getName();
		int idx = filename.indexOf("201");
		if (idx != -1) {
			filename = filename.substring(0, idx);
		}
		FileOutputStream out = new FileOutputStream(OUTPUT_PATH + filename + "-记账凭证-" + year + type + EXTENSION); 
        workbook.write(out);  
        workbook.close();
        out.close();  
		
	}

	private void initDataList(List<ExcelModel> list, int maxNum) {
		if (list != null && !list.isEmpty()) {
			int index = BEGIN_ROW;
			boolean cheked = false;
			Set<String> dateSet = new HashSet<String>();
			Map<String, Integer> summaryMap = new HashMap<String, Integer>();
			dataList = new ArrayList<ExcelModel>();
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
					
					if (!cheked && dateSet.contains(data.getColumnB())) {
						continue;
					}
					dateSet.add(data.getColumnB());
					if (!dataList.contains(data)) {
						index++;
						String key = data.getColumnF();
						if (summaryMap.get(key) != null) {
							Integer num = summaryMap.get(key);
							if (checkIfEnough(data.getColumnB().substring(0, 4), data.getColumnF(), num)) {
								continue;
							}
							summaryMap.put(key, num + 1);
						} else {
							summaryMap.put(key, 1);
						}
						dataList.add(data);
					}
				}
				cheked = true;
				time++;
			}
		}
		
		Collections.sort(dataList, new DateComparator());
	}

	private boolean checkIfEnough(String year, String columnD, Integer num) {
		Integer max = limitRecordNumBySummary.get(year) == null ? DEFAULT_RECORD_NUM_BY_SUMMARY : limitRecordNumBySummary.get(year);
		if (max > num) {
			max = limitRecordNumBySummary.get(columnD) == null ? DEFAULT_RECORD_NUM : limitRecordNumBySummary.get(columnD);
			
			if (max > num) {
				return false;
			}
		}
		
		
		return true;
	}

	private void initLargeRecords(List<ExcelModel> list, int maxNum) {
		if (list != null && !list.isEmpty()) {
			largeRecords = new ArrayList<Double>();
			List<Double> amountList = new ArrayList<Double>();
			for (ExcelModel data : list) {
				amountList.add(Math.abs(Double.valueOf(data.getColumnE())));
			}
			
			Collections.sort(amountList);
			Collections.reverse(amountList);
			largeRecords = amountList.subList(0, Math.min(maxNum * DEFAULT_FACTOR, amountList.size()));
		}
	}

	private int getMaxNum(String year) {
		for (String key : limitRecordNumBySummary.keySet()) {
			if (year.contains(key)) {
				return limitRecordNumBySummary.get(key);
			}
		}
		return DEFAULT_RECORD_NUM_BY_SUMMARY;
	}

	private boolean checkIfValidData(ExcelModel data) {
		
		if (!largeRecords.contains(Math.abs(Double.valueOf(data.getColumnE())))) {
			return false;
		}
		
//		if (limitRecordNumBySummary.get(data.getColumnF())  == null) {
//			return false;
//		}
		
		return true;
	}

	private List<ExcelModel> readExcelData(File file, String type, String year) throws FileNotFoundException, IOException {
		Workbook workbook = new HSSFWorkbook(new FileInputStream(file));
		Sheet sheet = workbook.getSheetAt(0);
		Row row;
		List<ExcelModel> list = new ArrayList<ExcelModel>();
		for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row != null) {
				
				if (!checkIfValidData(row, type, year)) {
					continue;
				}
				
				ExcelModel model = new ExcelModel();
				model.setColumnA(null);
				model.setColumnB(getCellValue(row.getCell(3)));
				model.setColumnC(getCellValue(row.getCell(2)));
				model.setColumnD(getCellValue(row.getCell(14)));
				model.setColumnE(getCellValue(row.getCell(18)));
				model.setColumnF(getCellValue(row.getCell(22)));
				model.setColumnG(null);
				
				list.add(model);
			}
		}
        workbook.close();

		return list;
	}

	private boolean checkIfValidData(Row row, String type, String year) {
		for (String filter : FILTER_COLUMN_SUMMARY) {
			if (filter.equals(getCellValue(row.getCell(22)).trim())) {
				return false;
			}
		}
		
		if (getCellValue(row.getCell(14)).startsWith("期末结转")) {
			return false;
		}
		
		if (getCellValue(row.getCell(14)).contains("转入")) {
			return false;
		}
		
		if (!getCellValue(row.getCell(22)).startsWith(type)) {
			return false;
		}
		
		if (!getCellValue(row.getCell(3)).startsWith(year)) {
			return false;
		}
		
		if ("0".equals(getCellValue(row.getCell(18)))) {
			return false;
		}
		
		return true;
	}

	private String getCellValue(Cell cell) {
		if (null != cell) {     
            switch (cell.getCellType()) {     
            case Cell.CELL_TYPE_NUMERIC: // 数字     
            	if (HSSFDateUtil.isCellDateFormatted(cell)) {
    				SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
    				return sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue())).toString();
    			}
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

	private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-m-d");
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

		private SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
		
		@Override
		public int compare(ExcelModel o1, ExcelModel o2) {
			String o1Date = "";
			String o2Date = "";
			try {
				o1Date = sdf.format(sdf.parse(o1.getColumnB()));
				o2Date = sdf.format(sdf.parse(o2.getColumnB()));
			} catch (ParseException e) {
				e.printStackTrace();
			}
			
			
			return o1Date.compareTo(o2Date);
		}
		
	}
}
