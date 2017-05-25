package com.report.rpt.source;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.StockRptDAO;
import com.report.model.StockSub2Model;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class StockSub2Source extends NDODailySalesCommonSource implements ReportSource{

	private StockRptDAO DAOobject;
	
	public StockSub2Source(StockRptDAO dAOobject){
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
		
//		CellStyle allBorderGrayAlignCenter = StyleUtils.allBorderGrayAlignCenter(sheet.getWorkbook().createCellStyle());
//		allBorderGrayAlignCenter.setFont(FontUtils.blackBold(sheet));
		
//		createBeforeHeaderForStock(sheet, startExcelRow-1, allBorderGrayAlignCenter);
		
    Cell headerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    headerColD.setCellStyle(allBorderYellowAlignCenter);
    headerColD.setCellValue("DoZoneName");
    
    Cell lastCol = createCommonHeader(sheet, ExcelUtils.getNextColumn(headerColD), criteria);	
//    createCommonHeaderForStock(sheet, lastCol, allBorderGrayAlignCenter, allBorderYellowAlignCenter);
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria) {
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		CellStyle allBorderAlignRight = StyleUtils.allBorderAlignRight(sheet.getWorkbook().createCellStyle());
		allBorderAlignRight.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));
		
		Cell doZoneNameCol  = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		List<Cell> dayCols = createDayCols(criteria, doZoneNameCol);
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
			
		List<StockSub2Model> stockSub2Models = getDAOobject().getSubReport2Data(criteria.getReportDate());
//		addTotalByDoZoneSubTotal1(stockSub2Models);
    for(StockSub2Model stockSub2Model:stockSub2Models){
    	
    	doZoneNameCol.setCellStyle(allBorder);
    	doZoneNameCol.setCellValue(stockSub2Model.getDoZoneName());
  		
    	setDayColsValue(dayCols, allBorderAlignRight, stockSub2Model);
    	
    	totalCol.setCellStyle(allBorderAlignRight);
    	totalCol.setCellValue(Utils.convertString2Int(stockSub2Model.getTotal()));
    	
    	lastMonthCol.setCellStyle(allBorderAlignRight);
    	lastMonthCol.setCellValue(Utils.convertString2Int(stockSub2Model.getLastMonth()));
    	
    	vsLastMthCol.setCellStyle(allBorderAlignRight);
    	vsLastMthCol.setCellType(CellType.FORMULA);
    	vsLastMthCol.setCellFormula(FormulasUtils.subtract(totalCol, lastMonthCol));
    	
    	// new
    	/*accumStockEndMonthCol.setCellStyle(allBorderAlignRight);
    	accumStockEndMonthCol.setCellValue(Utils.convertString2Int(stockSub2Model.getAccumStockEndLastMth()));
    	
    	accumNewWS.setCellStyle(allBorderAlignRight);
    	accumNewWS.setCellValue(Utils.convertString2Int(stockSub2Model.getAccumNewWS()));
    	
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
    	accumLastMonthStockEndMonthCol.setCellValue(Utils.convertString2Int(stockSub2Model.getAccumLastMonthStockEndLastMth()));
    	
    	accumLastMonthNewWS.setCellStyle(allBorderAlignRight);
    	accumLastMonthNewWS.setCellValue(Utils.convertString2Int(stockSub2Model.getAccumLastMonthNewWS()));
    	
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
    	if(stockSub2Model != stockSub2Models.get(stockSub2Models.size()-1)){
	    	doZoneNameCol = ExcelUtils.getNextRow(doZoneNameCol);
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
	
	/*
	public void addTotalByDoZoneSubTotal1(List<StockSub2Model> stockSub2Models){
		List<StockSub2Model> stockSub2ModelsTmp = new ArrayList<StockSub2Model>();
		StockSub2Model stockSub2Model;
		for(int i=0;i<stockSub2Models.size();i++){
			stockSub2Model = stockSub2Models.get(i);
			stockSub2ModelsTmp.add(stockSub2Model);
			if("04".equals(stockSub2Model.getZoneId())){
				stockSub2Model = new StockSub2Model();
				stockSub2Model.setDay1("0");
				stockSub2Model.setDay2("0");
				stockSub2Model.setDoZoneName("BKK & GB");
				stockSub2ModelsTmp.add(stockSub2Model);
			}
		}
		stockSub2Models.removeAll(stockSub2Models);
		stockSub2Models.addAll(stockSub2ModelsTmp);
		stockSub2ModelsTmp.removeAll(stockSub2ModelsTmp);
	}
	*/
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int... params) {
		Cell lastCol = createCommonFooter(sheet, startExcelColumn, startExcelRow, criteria, params[0]);
//		createCommonFooterForStock(sheet, lastCol, params[0]);
	}
}