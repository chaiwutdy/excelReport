package com.report.rpt.source;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.StockRptDAO;
import com.report.model.StockSub1Model;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class StockSub1Source extends NDODailySalesCommonSource implements ReportSource{

	private StockRptDAO DAOobject;
	
	public StockSub1Source(StockRptDAO dAOobject){
		DAOobject = dAOobject;
	}
	
	@Override
	public StockRptDAO getDAOobject() {
		return DAOobject;
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		CellStyle allBorderYellowAlignCenter = StyleUtils.allBorderYellowAlignCenter(sheet.getWorkbook().createCellStyle());
		allBorderYellowAlignCenter.setFont(FontUtils.blackBold(sheet));
		/*
		CellStyle allBorderGrayAlignCenter = StyleUtils.allBorderGrayAlignCenter(sheet.getWorkbook().createCellStyle());
		allBorderGrayAlignCenter.setFont(FontUtils.blackBold(sheet));
		
		createBeforeHeaderForStock(sheet, startExcelRow-1, allBorderGrayAlignCenter);
		*/
    Cell headerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    headerColD.setCellStyle(allBorderYellowAlignCenter);
    headerColD.setCellValue("Model");
    
    Cell lastCol = createCommonHeader(sheet, ExcelUtils.getNextColumn(headerColD), criteria);	
//    createCommonHeaderForStock(sheet, lastCol, allBorderGrayAlignCenter, allBorderYellowAlignCenter);
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
		/*Cell accumStockEndMonthCol = ExcelUtils.getNextColumn(vsLastMthCol);
		Cell accumNewWS = ExcelUtils.getNextColumn(accumStockEndMonthCol);
		Cell accumTotal1 = ExcelUtils.getNextColumn(accumNewWS);
		
		Cell accumWSCrossSaleInCol = ExcelUtils.getNextColumn(accumTotal1);
		Cell accumWSCrossSaleOutCol = ExcelUtils.getNextColumn(accumWSCrossSaleInCol);
		Cell accumTotal2 = ExcelUtils.getNextColumn(accumWSCrossSaleOutCol);
		
		Cell accumLastMonthStockEndMonthCol = ExcelUtils.getNextColumn(accumTotal2);
		Cell accumLastMonthNewWS = ExcelUtils.getNextColumn(accumLastMonthStockEndMonthCol);
		Cell accumLastMonthTotal1 = ExcelUtils.getNextColumn(accumLastMonthNewWS);
		
		Cell accumLastMonthWSCrossSaleInCol = ExcelUtils.getNextColumn(accumLastMonthTotal1);
		Cell accumLastMonthWSCrossSaleOutCol = ExcelUtils.getNextColumn(accumLastMonthWSCrossSaleInCol);
		Cell accumLastMonthTotal2 = ExcelUtils.getNextColumn(accumLastMonthWSCrossSaleOutCol);
		
		Cell accumGAPPreviousMonth = ExcelUtils.getNextColumn(accumLastMonthTotal2);
		Cell accumGAPThisMonth = ExcelUtils.getNextColumn(accumGAPPreviousMonth);*/
			
		List<StockSub1Model> stockSub1Models = getDAOobject().getSubReport1Data(criteria.getReportDate());
    for(StockSub1Model stockSub1Model:stockSub1Models){
    	
    	modelCol.setCellStyle(allBorder);
    	modelCol.setCellValue(stockSub1Model.getModel());
  		
    	setDayColsValue(dayCols, allBorderAlignRight, stockSub1Model);
    	
    	totalCol.setCellStyle(allBorderAlignRight);
    	totalCol.setCellValue(Utils.convertString2Int(stockSub1Model.getTotal()));
    	
    	lastMonthCol.setCellStyle(allBorderAlignRight);
    	lastMonthCol.setCellValue(Utils.convertString2Int(stockSub1Model.getLastMonth()));
    	
    	vsLastMthCol.setCellStyle(allBorderAlignRight);
    	vsLastMthCol.setCellType(CellType.FORMULA);
    	vsLastMthCol.setCellFormula(FormulasUtils.subtract(totalCol, lastMonthCol));
    	
    	// new
    	/*accumStockEndMonthCol.setCellStyle(allBorderAlignRight);
    	accumStockEndMonthCol.setCellValue(Utils.convertString2Int(stockSub1Model.getAccumStockEndLastMth()));
    	
    	accumNewWS.setCellStyle(allBorderAlignRight);
    	accumNewWS.setCellValue(Utils.convertString2Int(stockSub1Model.getAccumNewWS()));
    	
    	accumTotal1.setCellStyle(allBorderAlignRight);
    	accumTotal1.setCellType(CellType.FORMULA);
    	accumTotal1.setCellFormula(FormulasUtils.plus(accumStockEndMonthCol,accumNewWS));
    	
    	accumWSCrossSaleInCol.setCellStyle(allBorderAlignRight);
    	accumWSCrossSaleInCol.setCellType(CellType.FORMULA);
    	accumWSCrossSaleInCol.setCellFormula(getDiffReturn0WhenMinus(totalCol, accumTotal1));
    	
    	accumWSCrossSaleOutCol.setCellStyle(allBorderAlignRight);
    	accumWSCrossSaleOutCol.setCellType(CellType.FORMULA);
    	accumWSCrossSaleOutCol.setCellFormula(getDiffReturn0WhenMinus(accumTotal1, totalCol));
    	
    	accumTotal2.setCellStyle(allBorderAlignRight);
    	accumTotal2.setCellType(CellType.FORMULA);
    	accumTotal2.setCellFormula(FormulasUtils.subtract(accumWSCrossSaleInCol,accumWSCrossSaleOutCol));
    	
    	accumLastMonthStockEndMonthCol.setCellStyle(allBorderAlignRight);
    	accumLastMonthStockEndMonthCol.setCellValue(0);
    	accumLastMonthStockEndMonthCol.setCellValue(Utils.convertString2Int(stockSub1Model.getAccumLastMonthStockEndLastMth()));
    	
    	accumLastMonthNewWS.setCellStyle(allBorderAlignRight);
    	accumLastMonthNewWS.setCellValue(Utils.convertString2Int(stockSub1Model.getAccumLastMonthNewWS()));
    	
    	accumLastMonthTotal1.setCellStyle(allBorderAlignRight);
    	accumLastMonthTotal1.setCellType(CellType.FORMULA);
    	accumLastMonthTotal1.setCellFormula(FormulasUtils.plus(accumLastMonthStockEndMonthCol,accumLastMonthNewWS));
    	
    	accumLastMonthWSCrossSaleInCol.setCellStyle(allBorderAlignRight);
    	accumLastMonthWSCrossSaleInCol.setCellType(CellType.FORMULA);
    	accumLastMonthWSCrossSaleInCol.setCellFormula(getDiffReturn0WhenMinus(accumStockEndMonthCol, accumLastMonthTotal1));
    	
    	accumLastMonthWSCrossSaleOutCol.setCellStyle(allBorderAlignRight);
    	accumLastMonthWSCrossSaleOutCol.setCellType(CellType.FORMULA);
    	accumLastMonthWSCrossSaleOutCol.setCellFormula(getDiffReturn0WhenMinus(accumLastMonthTotal1, accumStockEndMonthCol));
    	
    	accumLastMonthTotal2.setCellStyle(allBorderAlignRight);
    	accumLastMonthTotal2.setCellType(CellType.FORMULA);
    	accumLastMonthTotal2.setCellFormula(FormulasUtils.subtract(accumLastMonthWSCrossSaleInCol,accumLastMonthWSCrossSaleOutCol));
    	
    	accumGAPPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumGAPPreviousMonth.setCellValue("");
    	
    	accumGAPThisMonth.setCellStyle(allBorderAlignRight);
    	accumGAPThisMonth.setCellValue("");*/
    	
    	//--------------------------- Create NextRow ---------------------------//
    	if(stockSub1Model != stockSub1Models.get(stockSub1Models.size()-1)){
	    	modelCol = ExcelUtils.getNextRow(modelCol);
	  		createDayColsNextRow(dayCols);
	  		totalCol = ExcelUtils.getNextRow(totalCol);
	  		lastMonthCol = ExcelUtils.getNextRow(lastMonthCol);
	  		vsLastMthCol = ExcelUtils.getNextRow(vsLastMthCol);
	  			  		
	  		
	  		// new 
	  		/*accumStockEndMonthCol = ExcelUtils.getNextRow(accumStockEndMonthCol);
	  		accumNewWS = ExcelUtils.getNextRow(accumNewWS);
	  		accumTotal1 = ExcelUtils.getNextRow(accumTotal1);
	  		accumWSCrossSaleInCol = ExcelUtils.getNextRow(accumWSCrossSaleInCol);
	  		accumWSCrossSaleOutCol = ExcelUtils.getNextRow(accumWSCrossSaleOutCol);
	  		accumTotal2 = ExcelUtils.getNextRow(accumTotal2);
	  		
	  		accumLastMonthStockEndMonthCol = ExcelUtils.getNextRow(accumLastMonthStockEndMonthCol);
	  		accumLastMonthNewWS = ExcelUtils.getNextRow(accumLastMonthNewWS);
	  		accumLastMonthTotal1 = ExcelUtils.getNextRow(accumLastMonthTotal1);
	  		accumLastMonthWSCrossSaleInCol = ExcelUtils.getNextRow(accumLastMonthWSCrossSaleInCol);
	  		accumLastMonthWSCrossSaleOutCol = ExcelUtils.getNextRow(accumLastMonthWSCrossSaleOutCol);
	  		accumLastMonthTotal2 = ExcelUtils.getNextRow(accumLastMonthTotal2);
	  		
	  		accumGAPPreviousMonth = ExcelUtils.getNextRow(accumGAPPreviousMonth);
	  		accumGAPThisMonth = ExcelUtils.getNextRow(accumGAPThisMonth);*/
    	}
    	
    }
	}
	
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int... params) {
		Cell lastCol = createCommonFooter(sheet, startExcelColumn, startExcelRow, criteria, params[0]);
//		createCommonFooterForStock(sheet, lastCol, params[0]);
	}
}