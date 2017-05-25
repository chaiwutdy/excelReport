package com.report.rpt.source;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.BackOrderRptDAO;
import com.report.model.BackOrderSub2Model;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class BackOrderSub2Source extends NDODailySalesCommonSource implements ReportSource{

private BackOrderRptDAO DAOobject;
	
	public BackOrderSub2Source(BackOrderRptDAO dAOobject){
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
    headerColD.setCellValue("DoZoneName");
    
    Cell lastCol = createCommonHeader(sheet, ExcelUtils.getNextColumn(headerColD), criteria);	
    createCommonHeaderForBackOrder(sheet, lastCol, allBorderGrayAlignCenter);
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
		Cell accumPreviousMonthCol = ExcelUtils.getNextColumn(vsLastMthCol);
		Cell accumThisMonth = ExcelUtils.getNextColumn(accumPreviousMonthCol);
		
		Cell accumLastMonthPreviousMonth = ExcelUtils.getNextColumn(accumThisMonth);
		Cell accumLastMonthThisMonth = ExcelUtils.getNextColumn(accumLastMonthPreviousMonth);
		
		Cell accumGAPPreviousMonth = ExcelUtils.getNextColumn(accumLastMonthThisMonth);
		Cell accumGAPThisMonth = ExcelUtils.getNextColumn(accumGAPPreviousMonth);

		List<BackOrderSub2Model> stockSub2Models = getDAOobject().getSubReport2Data(criteria.getReportDate());
//		addTotalByDoZoneSubTotal1(stockSub2Models);
    for(BackOrderSub2Model backOrderSub2Model:stockSub2Models){
    	
    	doZoneNameCol.setCellStyle(allBorder);
    	doZoneNameCol.setCellValue(backOrderSub2Model.getDoZoneName());
  		
    	setDayColsValue(dayCols, allBorderAlignRight, backOrderSub2Model);
    	
    	totalCol.setCellStyle(allBorderAlignRight);
    	totalCol.setCellValue(Utils.convertString2Int(backOrderSub2Model.getTotal()));
    	
    	lastMonthCol.setCellStyle(allBorderAlignRight);
    	lastMonthCol.setCellValue(Utils.convertString2Int(backOrderSub2Model.getLastMonth()));
    	
    	vsLastMthCol.setCellStyle(allBorderAlignRight);
    	vsLastMthCol.setCellType(CellType.FORMULA);
    	vsLastMthCol.setCellFormula(FormulasUtils.subtract(totalCol, lastMonthCol));
    	
    	// new
    	accumPreviousMonthCol.setCellStyle(allBorderAlignRight);
    	accumPreviousMonthCol.setCellValue(Utils.convertString2Int(backOrderSub2Model.getAccumPreviousMonth()));
    	
    	accumThisMonth.setCellStyle(allBorderAlignRight);
    	accumThisMonth.setCellValue(Utils.convertString2Int(backOrderSub2Model.getAccumThisMonth()));
    	
    	accumLastMonthPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumLastMonthPreviousMonth.setCellValue(Utils.convertString2Int(backOrderSub2Model.getAccumLastMonthPreviousMonth()));
    	
    	accumLastMonthThisMonth.setCellStyle(allBorderAlignRight);
    	accumLastMonthThisMonth.setCellValue(Utils.convertString2Int(backOrderSub2Model.getAccumLastMonthThisMonth()));
    	
    	accumGAPPreviousMonth.setCellStyle(allBorderAlignRight);
    	accumGAPPreviousMonth.setCellType(CellType.FORMULA);
    	accumGAPPreviousMonth.setCellFormula(FormulasUtils.subtract(accumPreviousMonthCol, accumLastMonthPreviousMonth));
    	
    	accumGAPThisMonth.setCellStyle(allBorderAlignRight);
    	accumGAPThisMonth.setCellType(CellType.FORMULA);
    	accumGAPThisMonth.setCellFormula(FormulasUtils.subtract(accumThisMonth, accumLastMonthThisMonth));
    	
    	//--------------------------- Create NextRow ---------------------------//
    	if(backOrderSub2Model != stockSub2Models.get(stockSub2Models.size()-1)){
	    	doZoneNameCol = ExcelUtils.getNextRow(doZoneNameCol);
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
	
	/*
	public void addTotalByDoZoneSubTotal1(List<BackOrderSub2Model> stockSub2Models){
		List<BackOrderSub2Model> stockSub2ModelsTmp = new ArrayList<BackOrderSub2Model>();
		BackOrderSub2Model backOrderSub2Model;
		for(int i=0;i<stockSub2Models.size();i++){
			backOrderSub2Model = stockSub2Models.get(i);
			stockSub2ModelsTmp.add(backOrderSub2Model);
			if("04".equals(backOrderSub2Model.getZoneId())){
				backOrderSub2Model = new BackOrderSub2Model();
				backOrderSub2Model.setDay1("0");
				backOrderSub2Model.setDay2("0");
				backOrderSub2Model.setDoZoneName("BKK & GB");
				stockSub2ModelsTmp.add(backOrderSub2Model);
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
		createCommonFooterForBackOrder(sheet, lastCol, params[0]);
	}
}