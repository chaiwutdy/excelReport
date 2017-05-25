package com.report.rpt.source;

import java.text.SimpleDateFormat;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.StyleUtils;

public class BackOrderMainSource extends NDODailySalesCommonSource implements ReportSource{

private String reportDetail;
	
	public BackOrderMainSource(String reportDetail){
		this.reportDetail = reportDetail;
	}
	
	public String getReportDetail() {
		return reportDetail;
	}
	
	@Override
	public Object getDAOobject() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		CellStyle allBorderGrayAlignCenter = StyleUtils.allBorderGrayAlignCenter(sheet.getWorkbook().createCellStyle());
		allBorderGrayAlignCenter.setFont(FontUtils.blackBold(sheet));
		
		Cell reportDate = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy",Locale.US);
    reportDate.setCellValue("Daily Stock remain Result as of "+sdf.format(criteria.getReportDate()));
    
    Cell report = ExcelUtils.getNextRow(reportDate);
    report.setCellValue("Total Back order remain (Previous & This month)");//get reportDetail

    Cell reportDetail = ExcelUtils.getNextRow(report);
    reportDetail.setCellValue(getReportDetail());
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void createFooter(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria, int... params) {
		// TODO Auto-generated method stub
		
	}
}
