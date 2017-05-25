package com.report.rpt.source;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.BackOrderRptDAO;
import com.report.model.BackOrderSub3Model;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class BackOrderSub3Source extends NDODailySalesCommonSource implements ReportSource{

private BackOrderRptDAO DAOobject;
	
	public BackOrderSub3Source(BackOrderRptDAO dAOobject){
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
    headerColD.setCellValue("DealerShortName");
		
    Cell lastCol = createCommonHeader(sheet, ExcelUtils.getNextColumn(headerColD), criteria);	
    createCommonHeaderForBackOrder(sheet, lastCol, allBorderGrayAlignCenter);
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		CellStyle allBorderAlignRight = StyleUtils.allBorderAlignRight(sheet.getWorkbook().createCellStyle());
		allBorderAlignRight.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));
		
		Cell zoneIdCol  = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		Cell areaIdCol  = ExcelUtils.getNextColumn(zoneIdCol);
		Cell dealerCodeCol  = ExcelUtils.getNextColumn(areaIdCol);
		Cell dealerShortNameCol  = ExcelUtils.getNextColumn(dealerCodeCol);
		List<Cell> dayCols = createDayCols(criteria, dealerShortNameCol);
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

		List<BackOrderSub3Model> stockSub3Models = getDAOobject().getSubReport3Data(criteria.getReportDate());
    for(BackOrderSub3Model backOrderSub3Model:stockSub3Models){
    	
    	zoneIdCol.setCellStyle(allBorder);
    	zoneIdCol.setCellValue(backOrderSub3Model.getZoneId());
    	
    	areaIdCol.setCellStyle(allBorder);
    	areaIdCol.setCellValue(backOrderSub3Model.getAreaId());
    	
    	dealerCodeCol.setCellStyle(allBorder);
    	dealerCodeCol.setCellValue(backOrderSub3Model.getDealerCode());
    	
    	dealerShortNameCol.setCellStyle(allBorder);
    	dealerShortNameCol.setCellValue(backOrderSub3Model.getDealerShortName());
  		
    	setDayColsValue(dayCols, allBorderAlignRight, backOrderSub3Model);
    	
    	totalCol.setCellStyle(allBorderAlignRight);
    	totalCol.setCellValue(Utils.convertString2Int(backOrderSub3Model.getTotal()));
    	
    	lastMonthCol.setCellStyle(allBorderAlignRight);
    	lastMonthCol.setCellValue(Utils.convertString2Int(backOrderSub3Model.getLastMonth()));
    	
    	vsLastMthCol.setCellStyle(allBorderAlignRight);
    	vsLastMthCol.setCellType(CellType.FORMULA);
    	vsLastMthCol.setCellFormula(FormulasUtils.subtract(totalCol, lastMonthCol));
    	
    	// new
    	accumPreviousMonthCol.setCellStyle(allBorderAlignRight);
    	accumPreviousMonthCol.setCellValue(Utils.convertString2Int(backOrderSub3Model.getAccumPreviousMonth()));
    	
    	accumThisMonth.setCellStyle(allBorderAlignRight);
    	accumThisMonth.setCellValue(Utils.convertString2Int(backOrderSub3Model.getAccumThisMonth()));
    	
    	accumLastMonthPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumLastMonthPreviousMonth.setCellValue(Utils.convertString2Int(backOrderSub3Model.getAccumLastMonthPreviousMonth()));
    	
    	accumLastMonthThisMonth.setCellStyle(allBorderAlignRight);
    	accumLastMonthThisMonth.setCellValue(Utils.convertString2Int(backOrderSub3Model.getAccumLastMonthThisMonth()));
    	
    	accumGAPPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumGAPPreviousMonth.setCellType(CellType.FORMULA);
    	accumGAPPreviousMonth.setCellFormula(FormulasUtils.subtract(accumPreviousMonthCol, accumLastMonthPreviousMonth));
    	
    	accumGAPThisMonth.setCellStyle(allBorderAlignRight);
    	accumGAPThisMonth.setCellType(CellType.FORMULA);
    	accumGAPThisMonth.setCellFormula(FormulasUtils.subtract(accumThisMonth, accumLastMonthThisMonth));
    	
    	//--------------------------- Create NextRow ---------------------------//
    	if(backOrderSub3Model != stockSub3Models.get(stockSub3Models.size()-1)){
	    	zoneIdCol = ExcelUtils.getNextRow(zoneIdCol);
	  		areaIdCol = ExcelUtils.getNextRow(areaIdCol);
	  		dealerCodeCol = ExcelUtils.getNextRow(dealerCodeCol);
	  		dealerShortNameCol = ExcelUtils.getNextRow(dealerShortNameCol);
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