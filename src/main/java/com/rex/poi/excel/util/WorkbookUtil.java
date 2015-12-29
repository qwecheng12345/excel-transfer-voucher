package com.rex.poi.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookUtil {
	public static final String EXCEL2003_EXT = "xls";
	public static final String EXCEL2007_EXT = "xlsx";
	
	private static SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
	private static DecimalFormat df = new DecimalFormat("###0.####");
	
	public static Workbook getWorkbook(File file) {
		String ext = FileUtil.getFileExtension(file);
		try {
			if (EXCEL2003_EXT.equalsIgnoreCase(ext)) {
				return new HSSFWorkbook(new FileInputStream(file));
			} else if (EXCEL2007_EXT.equalsIgnoreCase(ext)) {
				return new XSSFWorkbook(new FileInputStream(file));
			}
		} catch (Exception e) {
			return null;
		}
		return null;
	}

	public static Workbook getSimpleTemplate(String filePath) {
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filePath));
			return new HSSFWorkbook(fs);
			
		} catch (Exception e) {
			return null;
		}
	}
	
	public static String getCellValue(Cell cell) {
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
