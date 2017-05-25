package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.DummyRptDAO;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

/**
 * This class is a ReportSource Example .
 */
public class DummyRptSource implements ReportSource{

	@Override
	public DummyRptDAO getDAOobject() {
		return (DummyRptDAO)Utils.APP_CONTEXT.getBean("dummyRptDAO");
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria) {
		Cell reportDate = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    reportDate.setCellValue("DUMMY REPORT AS "+criteria.getReportDate());
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria) {
		String reportData = getDAOobject().getReportData();
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		
		Cell bodyColB =  ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		bodyColB.setCellStyle(allBorder);
		bodyColB.setCellValue("DUMMY B");
		
		Cell bodyColC =  ExcelUtils.getNextColumn(bodyColB);
		bodyColC.setCellStyle(allBorder);
		bodyColC.setCellValue(reportData);
		
		sheet.autoSizeColumn(bodyColB.getColumnIndex());
		sheet.autoSizeColumn(bodyColC.getColumnIndex());
		
		//
		Cell bodyColD =  ExcelUtils.getNextColumn(bodyColC);
		bodyColD.setCellStyle(allBorder);
		bodyColD.setCellValue(1);
		
		Cell bodyColDNextRow =  ExcelUtils.getNextRow(bodyColD);
		bodyColDNextRow.setCellStyle(allBorder);
		bodyColDNextRow.setCellValue(1);
		
		
		//----- TEST MERAGE COLUMN -----//
		Cell cellE1 =  ExcelUtils.getCell("E1",sheet);
		cellE1.setCellValue("TEST MERAGE COLUMN");
		
		Cell cellG1 =  ExcelUtils.getCell("G1",sheet);
		ExcelUtils.mergeColumn(sheet, cellE1, cellG1);
		
		//----- TEST MERAGE ROW -----//
		Cell cellE4 =  ExcelUtils.getCell("E4",sheet);
		cellE4.setCellValue("TEST MERAGE ROW");
		
		Cell cellE8 =  ExcelUtils.getCell("E8",sheet);
		ExcelUtils.mergeRow(sheet, cellE4, cellE8);
		
		
		//----- TEST MERAGE CELL -----//
		Cell cellE9 =  ExcelUtils.getCell("E9",sheet);
		cellE9.setCellValue("TEST MERAGE CELL");
		
		Cell cellG10 =  ExcelUtils.getCell("G10",sheet);
		ExcelUtils.mergeCell(sheet, cellE9, cellG10);
		
	}

	@Override
	public void createFooter(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria, int... params) {
		Cell bodyColB =  ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		bodyColB.setCellValue("Total");
		
		Cell bodyColC =  ExcelUtils.getNextColumn(bodyColB);
		Cell bodyColD =  ExcelUtils.getNextColumn(bodyColC);
//		bodyColD.setCellFormula("SUM(D"+params[0]+":D"+(startExcelRow-1)+")");
		bodyColD.setCellFormula(FormulasUtils.sum(bodyColD, params[0], startExcelRow-1));
	}
	
}
