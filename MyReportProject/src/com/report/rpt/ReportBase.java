package com.report.rpt;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public interface ReportBase {
	public String getFilePath();
	public Object getDAOobject();
	public void createHeader(Sheet sheet);
	public void createBody(Sheet sheet);
	public Workbook getWorkbook();
}
