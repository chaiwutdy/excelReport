package com.report.rpt;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.dao.DummyRptDAO;
import com.report.util.DateUtils;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class DummyRpt implements ReportBase{

	@Override
	public String getFilePath() {
		return Utils.getProperties("app.path.dummyRpt");
	}

	@Override
	public DummyRptDAO getDAOobject() {
		return (DummyRptDAO)Utils.APP_CONTEXT.getBean("dummyRptDAO");
	}

	@Override
	public void createHeader(Sheet sheet) {
		Cell reportDate = ExcelUtils.getCell("A1",sheet);
    reportDate.setCellValue("DUMMY REPORT AS "+DateUtils.getCurrentDate("dd-MM-yyyy"));
	}

	@Override
	public void createBody(Sheet sheet) {
		String reportData = getDAOobject().getReportData();
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		
		Cell bodyColB =  ExcelUtils.getCell("B2", sheet);
		bodyColB = ExcelUtils.getCell("B2", sheet);
		bodyColB.setCellStyle(allBorder);
		bodyColB.setCellValue("DUMMY B2");
		
		Cell bodyColC =  ExcelUtils.getCell("C2", sheet);
		bodyColC = ExcelUtils.getCell("C2", sheet);
		bodyColC.setCellStyle(allBorder);
		bodyColC.setCellValue(reportData);
		
		sheet.autoSizeColumn(bodyColB.getColumnIndex());
		sheet.autoSizeColumn(bodyColC.getColumnIndex());
	}

	@Override
	public Workbook getWorkbook() {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("SheetDummy");
    createHeader(sheet);
    createBody(sheet);
		return book;
	}

}
