package com.report.generator;

import org.apache.poi.ss.usermodel.Workbook;

public interface ReportGenerator {
	public String getFilePath();
	public Workbook getWorkbook();
}
