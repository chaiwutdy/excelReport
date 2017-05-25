package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;

public class IvnAdvRptSitMainSource extends IvnAdvRptSitCommonSource implements ReportSource{

	@Override
	public Object getDAOobject() {
		return null;
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria) {
		Cell reportDate = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    reportDate.setCellValue("INVOICE ADVANCE SITUATION AS "+criteria.getReportDate());
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria) {
	}
	
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria, int... params) {
	}
	
	
}
