package com.report.rpt.source;

import java.text.SimpleDateFormat;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;

public class StockMainSource extends NDODailySalesCommonSource implements ReportSource{

	private String reportDetail;
	
	public StockMainSource(String reportDetail){
		this.reportDetail = reportDetail;
	}
	
	public String getReportDetail() {
		return reportDetail;
	}

	@Override
	public Object getDAOobject() {
		return null;
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		Cell reportDate = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy",Locale.US);
    reportDate.setCellValue("Daily Stock remain Result as of "+sdf.format(criteria.getReportDate()));
    
    Cell report = ExcelUtils.getNextRow(reportDate);
    report.setCellValue(getReportDetail());
		
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
	}

	@Override
	public void createFooter(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria, int... params) {
	}
}
