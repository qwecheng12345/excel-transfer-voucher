package com.rex.poi.excel.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.rex.poi.excel.model.SourceModel;
import com.rex.poi.excel.model.TargetModel;
import com.rex.poi.excel.util.FileUtil;

public class FindSubject2 {
	private static final String PATH = "E:\\Rex\\01_Project\\60_TmpCode\\99_Minyee\\input\\20151229\\source.xls";
	private static final String O_PATH = "E:\\Rex\\01_Project\\60_TmpCode\\99_Minyee\\input\\20151229\\Book1.xls";
	
	public static List<SourceModel> sourceList = new ArrayList<SourceModel>();
	public static List<TargetModel> targetList = new ArrayList<TargetModel>();
	
	public static void main(String[] args) throws Exception {
		getSourceDataList();
		System.out.println(sourceList.size());
		setTargetData();
	}
	
	private static void getSourceDataList() throws Exception{
		File file = new File(PATH);
		Workbook workbook = new HSSFWorkbook(new FileInputStream(file));
		for (int s = 1; s < 7; s++) {
			if (s == 3 || s == 6) {
				continue;
			}
			Sheet sheet = workbook.getSheetAt(s);
			Row row;
			System.out.println(sheet.getPhysicalNumberOfRows());
			for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
				row = sheet.getRow(i);
				if (row != null) {
					if (!checkIfValidData(row)) {
						continue;
					}
					SourceModel model = new SourceModel();
					model.setColumnA(getCellValue(row.getCell(3)));
					model.setColumnB(getCellValue(row.getCell(2)));
					model.setColumnC(getCellValue(row.getCell(10)));
					model.setColumnD(getCellValue(row.getCell(13)));
					model.setColumnE(getCellValue(row.getCell(17)));
					model.setChecked(false);
					sourceList.add(model);
				}
			}
		}
		workbook.close();
	}

	private static void setTargetSubject(TargetModel model) {
		for (SourceModel source : sourceList) {
			if (source.isChecked()) {
				continue;
			}
			
			if (source.getColumnA().equals(model.getColumnA())
					&& source.getColumnB().equals(model.getColumnB())
					&& source.getColumnC().contains(model.getColumnC())
					&& source.getColumnD().equals(model.getColumnD())) {
				source.setChecked(true);
				model.setColumnE(source.getColumnE());
				break;
			}
		}
	}

	private static boolean checkIfValidData(Row row) {
		try {
			if (row.getCell(2) == null || "".equals(row.getCell(2))
					|| row.getCell(3) == null || "".equals(row.getCell(3))
					|| row.getCell(10) == null || "".equals(row.getCell(10))
					|| row.getCell(13) == null || "".equals(row.getCell(13))) {
				return false;			
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	private static void setTargetData() throws Exception {
		File template = new File(O_PATH);  
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(template));
        Workbook workbook = new HSSFWorkbook(fs);
        Sheet sheet = workbook.getSheetAt(0);
        Row row;
        Cell cell;
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row != null) {
				if (!checkIfValidTargetData(row)) {
					continue;
				}
				TargetModel model = new TargetModel();
				model.setColumnA(getCellValue(row.getCell(0)));
				if (getCellValue(row.getCell(1)) != null && getCellValue(row.getCell(1)).length() > 2) {
					if (FileUtil.beginWithDigit(getCellValue(row.getCell(1)))) {
						model.setColumnB(getCellValue(row.getCell(1)));
					} else {
						model.setColumnB(getCellValue(row.getCell(1)).substring(2));
					}
				} else {
					model.setColumnB(getCellValue(row.getCell(1)));
				}
				model.setColumnC(getCellValue(row.getCell(2)));
				model.setColumnD(getCellValue(row.getCell(3)));
				setTargetSubject(model);
				
				cell = row.createCell(4);
				cell.setCellValue(model.getColumnE());
			}
		}
        
        FileOutputStream out = new FileOutputStream(O_PATH);  
        workbook.write(out);  
        workbook.close();
        out.close(); 
	}
	
	private static boolean checkIfValidTargetData(Row row) {
		try {
			if (row.getCell(0) == null || "".equals(row.getCell(0))
					|| row.getCell(1) == null || "".equals(row.getCell(1))
					|| row.getCell(2) == null || "".equals(row.getCell(2))
					|| row.getCell(3) == null || "".equals(row.getCell(3))) {
				return false;			
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	private static String getCellValue(Cell cell) {
		if (null != cell) {     
            switch (cell.getCellType()) {     
            case Cell.CELL_TYPE_NUMERIC: // Êý×Ö     
            	if (HSSFDateUtil.isCellDateFormatted(cell)) {
    				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-M-d");
    				return sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue())).toString();
    			}
            	return formatVal(cell.getNumericCellValue());     
            case Cell.CELL_TYPE_STRING: // ×Ö·û´®     
            	return formatVal(cell.getStringCellValue());     
            case Cell.CELL_TYPE_BOOLEAN: // Boolean     
            	return formatVal(cell.getBooleanCellValue());     
            case Cell.CELL_TYPE_FORMULA: // ¹«Ê½     
            	return cell.getCellFormula();
            case Cell.CELL_TYPE_BLANK: // ¿ÕÖµ     
            	return "";
            case Cell.CELL_TYPE_ERROR: // ¹ÊÕÏ     
                break;     
            default:     
                break;     
            }     
        }
		return null;  
	}

	private static SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
	private static DecimalFormat df = new DecimalFormat("###0.####");
	private static String formatVal(Object val) {
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
	
	private static boolean isNotEmpty(Object object) {
		if (object != null) {
			if (object instanceof String) {
				return (!"".equals((String) object));
			}
			return true;
		}
		return false;
	}

}
