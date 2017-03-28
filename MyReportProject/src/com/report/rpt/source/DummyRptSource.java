package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.DummyRptDAO;
import com.report.util.DateUtils;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class DummyRptSource implements ReportSource{

	@Override
	public DummyRptDAO getDAOobject() {
		return (DummyRptDAO)Utils.APP_CONTEXT.getBean("dummyRptDAO");
	}

	@Override
	public void createHeader(Sheet sheet, int startRowNum) {
		Cell reportDate = ExcelUtils.getCell("A"+startRowNum, sheet);
    reportDate.setCellValue("DUMMY REPORT AS "+DateUtils.getCurrentDate("dd-MM-yyyy"));
	}

	@Override
	public void createBody(Sheet sheet, int startRowNum) {
		String reportData = getDAOobject().getReportData();
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		
		Cell bodyColB =  ExcelUtils.getCell("B"+startRowNum, sheet);
		bodyColB.setCellStyle(allBorder);
		bodyColB.setCellValue("DUMMY B");
		
		Cell bodyColC =  ExcelUtils.getCell("C"+startRowNum, sheet);
		bodyColC.setCellStyle(allBorder);
		bodyColC.setCellValue(reportData);
		
		sheet.autoSizeColumn(bodyColB.getColumnIndex());
		sheet.autoSizeColumn(bodyColC.getColumnIndex());
	}

	@Override
	public void createFooter(Sheet sheet, int startRowNum) {
		Cell bodyColB =  ExcelUtils.getCell("B"+startRowNum, sheet);
		bodyColB.setCellValue("Footer");
	}
	
}
