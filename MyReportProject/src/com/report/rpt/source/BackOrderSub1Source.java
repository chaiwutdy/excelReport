package com.report.rpt.source;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.BackOrderRptDAO;
import com.report.model.BackOrderSub1Model;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class BackOrderSub1Source extends NDODailySalesCommonSource implements ReportSource{

	private BackOrderRptDAO DAOobject;
	
	public BackOrderSub1Source(BackOrderRptDAO dAOobject){
		DAOobject = dAOobject;
	}
	
	@Override
	public BackOrderRptDAO getDAOobject() {
		return DAOobject;
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		CellStyle allBorderYellowAlignCenter = StyleUtils.allBorderYellowAlignCenter(sheet.getWorkbook().createCellStyle());
		allBorderYellowAlignCenter.setFont(FontUtils.blackBold(sheet));
		
		CellStyle allBorderGrayAlignCenter = StyleUtils.allBorderGrayAlignCenter(sheet.getWorkbook().createCellStyle());
		allBorderGrayAlignCenter.setFont(FontUtils.blackBold(sheet));
		
		createBeforeHeaderForBackOrder(sheet, startExcelRow-1, allBorderGrayAlignCenter);
		
    Cell headerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    headerColD.setCellStyle(allBorderYellowAlignCenter);
    headerColD.setCellValue("Model");
    
    Cell lastCol = createCommonHeader(sheet, ExcelUtils.getNextColumn(headerColD), criteria);	
    createCommonHeaderForBackOrder(sheet, lastCol, allBorderGrayAlignCenter);
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		CellStyle allBorderAlignRight = StyleUtils.allBorderAlignRight(sheet.getWorkbook().createCellStyle());
		allBorderAlignRight.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));
		
		Cell modelCol  = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		List<Cell> dayCols = createDayCols(criteria, modelCol);
		Cell totalCol = ExcelUtils.getNextColumn(dayCols.get(dayCols.size()-1));
		Cell lastMonthCol = ExcelUtils.getNextColumn(totalCol);
		Cell vsLastMthCol = ExcelUtils.getNextColumn(lastMonthCol);

		//new
		Cell accumPreviousMonthCol = ExcelUtils.getNextColumn(vsLastMthCol);
		Cell accumThisMonth = ExcelUtils.getNextColumn(accumPreviousMonthCol);
		
		Cell accumLastMonthPreviousMonth = ExcelUtils.getNextColumn(accumThisMonth);
		Cell accumLastMonthThisMonth = ExcelUtils.getNextColumn(accumLastMonthPreviousMonth);
		
		Cell accumGAPPreviousMonth = ExcelUtils.getNextColumn(accumLastMonthThisMonth);
		Cell accumGAPThisMonth = ExcelUtils.getNextColumn(accumGAPPreviousMonth);
		
		List<BackOrderSub1Model> stockSub1Models = getDAOobject().getSubReport1Data(criteria.getReportDate());
    for(BackOrderSub1Model backOrderSub1Model:stockSub1Models){
    	
    	modelCol.setCellStyle(allBorder);
    	modelCol.setCellValue(backOrderSub1Model.getModel());
  		
    	setDayColsValue(dayCols, allBorderAlignRight, backOrderSub1Model);
    	
    	totalCol.setCellStyle(allBorderAlignRight);
    	totalCol.setCellValue(Utils.convertString2Int(backOrderSub1Model.getTotal()));
    	
    	lastMonthCol.setCellStyle(allBorderAlignRight);
    	lastMonthCol.setCellValue(Utils.convertString2Int(backOrderSub1Model.getLastMonth()));
    	
    	vsLastMthCol.setCellStyle(allBorderAlignRight);
    	vsLastMthCol.setCellType(CellType.FORMULA);
    	vsLastMthCol.setCellFormula(FormulasUtils.subtract(totalCol, lastMonthCol));
    	
    	// new
    	accumPreviousMonthCol.setCellStyle(allBorderAlignRight);
    	accumPreviousMonthCol.setCellValue(Utils.convertString2Int(backOrderSub1Model.getAccumPreviousMonth()));
    	
    	accumThisMonth.setCellStyle(allBorderAlignRight);
    	accumThisMonth.setCellValue(Utils.convertString2Int(backOrderSub1Model.getAccumThisMonth()));
    	
    	accumLastMonthPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumLastMonthPreviousMonth.setCellValue(Utils.convertString2Int(backOrderSub1Model.getAccumLastMonthPreviousMonth()));
    	
    	accumLastMonthThisMonth.setCellStyle(allBorderAlignRight);
    	accumLastMonthThisMonth.setCellValue(Utils.convertString2Int(backOrderSub1Model.getAccumLastMonthThisMonth()));
    	
    	accumGAPPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumGAPPreviousMonth.setCellType(CellType.FORMULA);
    	accumGAPPreviousMonth.setCellFormula(FormulasUtils.subtract(accumPreviousMonthCol, accumLastMonthPreviousMonth));
    	
    	accumGAPThisMonth.setCellStyle(allBorderAlignRight);
    	accumGAPThisMonth.setCellType(CellType.FORMULA);
    	accumGAPThisMonth.setCellFormula(FormulasUtils.subtract(accumThisMonth, accumLastMonthThisMonth));
    	
    	//--------------------------- Create NextRow ---------------------------//
    	if(backOrderSub1Model != stockSub1Models.get(stockSub1Models.size()-1)){
	    	modelCol = ExcelUtils.getNextRow(modelCol);
	  		createDayColsNextRow(dayCols);
	  		totalCol = ExcelUtils.getNextRow(totalCol);
	  		lastMonthCol = ExcelUtils.getNextRow(lastMonthCol);
	  		vsLastMthCol = ExcelUtils.getNextRow(vsLastMthCol);
	  		
	  		// new 
	  		accumPreviousMonthCol = ExcelUtils.getNextRow(accumPreviousMonthCol);
	  		accumThisMonth = ExcelUtils.getNextRow(accumThisMonth);
	  		accumLastMonthPreviousMonth = ExcelUtils.getNextRow(accumLastMonthPreviousMonth);
	  		accumLastMonthThisMonth = ExcelUtils.getNextRow(accumLastMonthThisMonth);
	  		accumGAPPreviousMonth = ExcelUtils.getNextRow(accumGAPPreviousMonth);
	  		accumGAPThisMonth = ExcelUtils.getNextRow(accumGAPThisMonth);
    	}
    	
    }
	}
	
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int... params) {
		Cell lastCol = createCommonFooter(sheet, startExcelColumn, startExcelRow, criteria, params[0]);
		createCommonFooterForBackOrder(sheet, lastCol, params[0]);
	}
}