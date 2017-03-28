package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Sheet;

public interface ReportSource {
	public Object getDAOobject();
	public void createHeader(Sheet sheet, int startRowNum);
	public void createBody(Sheet sheet, int startRowNum);
	public void createFooter(Sheet sheet, int startRowNum);
}
