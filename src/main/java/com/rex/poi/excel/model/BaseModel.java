package com.rex.poi.excel.model;

public abstract class BaseModel {
	private String columnA;		// Date(日期, yyyy-MM-dd)
	private String columnB;		// Voucher(凭证号, A-B)
	private String columnC;		// Summary(摘要)
	private String columnD;		// Amount(金额, ###0.####)
	private String columnE;		// Subject(科目)
	public String getColumnA() {
		return columnA;
	}
	public void setColumnA(String columnA) {
		this.columnA = columnA;
	}
	public String getColumnB() {
		return columnB;
	}
	public void setColumnB(String columnB) {
		this.columnB = columnB;
	}
	public String getColumnC() {
		return columnC;
	}
	public void setColumnC(String columnC) {
		this.columnC = columnC;
	}
	public String getColumnD() {
		return columnD;
	}
	public void setColumnD(String columnD) {
		this.columnD = columnD;
	}
	public String getColumnE() {
		return columnE;
	}
	public void setColumnE(String columnE) {
		this.columnE = columnE;
	}
}
