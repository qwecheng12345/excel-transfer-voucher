package com.rex.poi.excel.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.rex.poi.excel.model.ExcelModel;
import com.rex.poi.excel.model.MonthCode;
import com.rex.poi.excel.util.FileUtil;
import com.rex.poi.excel.util.RandomUtil;
import com.rex.poi.excel.util.WorkbookUtil;

public class SimpleExcelServiceImpl implements ExcelService {
	
	private static Logger logger = Logger.getLogger(SimpleExcelServiceImpl.class);
	
	private static final String[] FILTER_COLUMN_E = {"期末结转"};
	private static final Integer DEFAULT_MAX_NUM = 80;
	
	private static final Integer BEGIN_ROW = 8;

	@Override
	public List<File> getFiles(String path) {
		List<File> fileList = FileUtil.getFiles(path);
		return fileList;
	}

	@Override
	public List<ExcelModel> getDatas(File file) {
		Workbook workbook = WorkbookUtil.getWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);
		Row row;
		List<ExcelModel> list = new ArrayList<ExcelModel>();
		for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row != null) {
				ExcelModel model = new ExcelModel();
				try {
					model.setColumnA(WorkbookUtil.getCellValue(row.getCell(0)));
					model.setColumnB(WorkbookUtil.getCellValue(row.getCell(1)));
					model.setColumnC(WorkbookUtil.getCellValue(row.getCell(2)));
					model.setColumnD(WorkbookUtil.getCellValue(row.getCell(3)));
					model.setColumnE(WorkbookUtil.getCellValue(row.getCell(4)));
					model.setColumnF(WorkbookUtil.getCellValue(row.getCell(5)));
					model.setColumnG(WorkbookUtil.getCellValue(row.getCell(6)));
					
					list.add(model);
				} catch (Exception e) {
					if (logger.isInfoEnabled()) {
						logger.error("Missing data from " + file.getName() + ", line " + row.getRowNum());
					}
				}
			}
		}
		return list;
	}

	@Override
	public Workbook dealWithDatas(List<ExcelModel> datas,File source, String template, boolean limited) {
		Workbook workbook = WorkbookUtil.getSimpleTemplate(template);
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
		
		CellStyle cellStyle3 = workbook.createCellStyle();
		font.setFontHeightInPoints((short) 11);
		cellStyle3.setFont(font);
		cellStyle3.setAlignment(CellStyle.ALIGN_GENERAL);
		cellStyle3.setVerticalAlignment(CellStyle.VERTICAL_BOTTOM);
		cellStyle3.setWrapText(false);
		cellStyle3.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle3.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle3.setBorderRight(CellStyle.BORDER_THIN);
		cellStyle3.setBorderTop(CellStyle.BORDER_THIN);
		cellStyle3.setDataFormat(workbook.createDataFormat().getFormat("yyyy/m/d"));
		
		if (limited) {
			datas = getLimitedDatas(datas);
		}
		
		if (datas != null && !datas.isEmpty()) {
			String year = FileUtil.getFileYear(source.getName());
			for (ExcelModel data : datas) {
				if (!checkDataIfValid(data)) {
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
				row.getCell(1).setCellStyle(cellStyle3);
				row.getCell(2).setCellStyle(cellStyle2);
				row.getCell(3).setCellStyle(cellStyle2);
				row.getCell(4).setCellStyle(cellStyle2);
				row.getCell(5).setCellStyle(cellStyle2);
				row.getCell(6).setCellStyle(cellStyle2);
				row.getCell(7).setCellStyle(cellStyle2);
				row.getCell(8).setCellStyle(cellStyle2);
			}
		}
		
		return workbook;
	}

	private List<ExcelModel> getLimitedDatas(List<ExcelModel> datas) {
		List<ExcelModel> result = new ArrayList<ExcelModel>();
//		List<Double> amountRange = initAmountRange(datas);
		if (datas != null && !datas.isEmpty()) {
			Map<MonthCode, List<Double>> dataMap = RandomUtil.getDataDescOrderPerMonth(datas);
			// TODO check if empty
			int monthNum = dataMap.keySet().size();
			int numPerMonth = DEFAULT_MAX_NUM / monthNum + 1;
			
			int index = BEGIN_ROW;
			boolean cheked = false;
			Set<String> monthSet = new HashSet<String>();
			int time = 0;
			while (DEFAULT_MAX_NUM > result.size() && time < 3) {
				for (ExcelModel data : datas) {
					if (index > BEGIN_ROW + DEFAULT_MAX_NUM - 1) {
						time = 3;
						break;
					}
					
					int month = Integer.valueOf(data.getColumnA());
					if (!checkIfValidData(dataMap.get(month), data, numPerMonth)) {
						continue;
					}
					
					if (!result.contains(data)) {
						index++;
						result.add(data);
					}
				}
				cheked = true;
				time++;
			}
		}
		
		Collections.sort(result, new DateComparator());
		return result;
	}

	private boolean checkIfValidData(List<Double> list, ExcelModel data,
			int numPerMonth) {
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
		if (!list.subList(0, numPerMonth).contains(amount)) {
			return false;
		}
		return true;
	}

	private List<Double> initAmountRange(List<ExcelModel> datas) {
		List<Double> amountList = new ArrayList<Double>();
		for (ExcelModel data : datas) {
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
		return amountList.subList(0, Math.min(DEFAULT_MAX_NUM * 2, amountList.size()));
	}

	private boolean checkDataIfValid(ExcelModel data) {
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
			return false;
		}
		
		return true;
	}

	@Override
	public void writeDataToExcel(Workbook workbook, String fileName, String outputPath) {
		FileOutputStream out = null;
		try {
			File file = new File(outputPath + "记账凭证-" + fileName);
			if (file.canWrite()) {
				out = new FileOutputStream(outputPath + "记账凭证-" + fileName);  
				workbook.write(out);  
			}
		} catch (Exception e) {
			
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {
				
			}
			try {
				out.close();
			} catch (IOException e) {
				
			}
		}
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
