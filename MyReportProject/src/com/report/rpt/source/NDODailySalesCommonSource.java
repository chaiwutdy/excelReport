package com.report.rpt.source;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import com.report.rpt.criteria.ReportCriteria;
import com.report.util.DateUtils;
import com.report.util.ExcelUtils;
import com.report.util.FontUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public abstract class NDODailySalesCommonSource {

//---------------------------------------------------- For Back Order ----------------------------------------------------//
	protected void createBeforeHeaderForBackOrder(Sheet sheet, int startExcelRow, CellStyle allBorderGrayAlignCenter){
		//------------------------------------- Accum Back Order -----------------------------------------//
		
    
    Cell accumBackOrder1H = ExcelUtils.getCell("AM", startExcelRow, sheet);
    accumBackOrder1H.setCellStyle(allBorderGrayAlignCenter);
    accumBackOrder1H.setCellValue("Accum Back Order");
    
    Cell accumBackOrder2H = ExcelUtils.getNextColumn(accumBackOrder1H);
    accumBackOrder2H.setCellStyle(allBorderGrayAlignCenter);
    ExcelUtils.mergeColumn(sheet, accumBackOrder1H, accumBackOrder2H);
    
    //------------------------------------- Accum Back Order Last month -----------------------------------------//
    allBorderGrayAlignCenter.setWrapText(true);
    Cell accumBackOrderLastMonth1H = ExcelUtils.getNextColumn(accumBackOrder2H);
    accumBackOrderLastMonth1H.setCellStyle(allBorderGrayAlignCenter);
    accumBackOrderLastMonth1H.setCellValue("Accum Back Order Last month");
    
    Cell accumBackOrderLastMonth2H = ExcelUtils.getNextColumn(accumBackOrderLastMonth1H);
    accumBackOrderLastMonth2H.setCellStyle(allBorderGrayAlignCenter);
    ExcelUtils.mergeColumn(sheet, accumBackOrderLastMonth1H, accumBackOrderLastMonth2H);
    
    //------------------------------------- Accum Back Order Last month -----------------------------------------//
    Cell accumGAPBackOrder1H = ExcelUtils.getNextColumn(accumBackOrderLastMonth2H);
    accumGAPBackOrder1H.setCellStyle(allBorderGrayAlignCenter);
    accumGAPBackOrder1H.setCellValue("Accum GAP Back Order");
    
    Cell accumGAPBackOrder2H = ExcelUtils.getNextColumn(accumGAPBackOrder1H);
    accumGAPBackOrder2H.setCellStyle(allBorderGrayAlignCenter);
    ExcelUtils.mergeColumn(sheet, accumGAPBackOrder1H, accumGAPBackOrder2H);
    //------------------------------------------------------------------------------------------------------------//
    
    Cell dummyCell = ExcelUtils.getNextColumn(accumGAPBackOrder2H);
		dummyCell.setCellStyle(allBorderGrayAlignCenter);
		dummyCell.setCellValue("use for new line");
		sheet.setColumnHidden(dummyCell.getColumnIndex(), true);
	}
	
	protected void createCommonHeaderForBackOrder(Sheet sheet, Cell beginCell, CellStyle allBorderGrayAlignCenter){
		Cell accumBackOrder1 = ExcelUtils.getNextColumn(beginCell);
    accumBackOrder1.setCellStyle(allBorderGrayAlignCenter);
    accumBackOrder1.setCellValue("Previous month");    
    
    Cell accumBackOrder2 = ExcelUtils.getNextColumn(accumBackOrder1);
    accumBackOrder2.setCellStyle(allBorderGrayAlignCenter);
    accumBackOrder2.setCellValue("This month");
    
    Cell accumBackOrderLastMonth1 = ExcelUtils.getNextColumn(accumBackOrder2);
    accumBackOrderLastMonth1.setCellStyle(allBorderGrayAlignCenter);
    accumBackOrderLastMonth1.setCellValue("Previous month");   
    
    Cell accumBackOrderLastMonth2 = ExcelUtils.getNextColumn(accumBackOrderLastMonth1);
    accumBackOrderLastMonth2.setCellStyle(allBorderGrayAlignCenter);
    accumBackOrderLastMonth2.setCellValue("This month");
    
    Cell accumGAPBackOrder1 = ExcelUtils.getNextColumn(accumBackOrderLastMonth2);
    accumGAPBackOrder1.setCellStyle(allBorderGrayAlignCenter);
    accumGAPBackOrder1.setCellValue("Previous month");   
    
    Cell accumGAPBackOrder2 = ExcelUtils.getNextColumn(accumGAPBackOrder1);
    accumGAPBackOrder2.setCellStyle(allBorderGrayAlignCenter);
    accumGAPBackOrder2.setCellValue("This month");	
	}

//---------------------------------------------------- For Stock ----------------------------------------------------//
	protected void createBeforeHeaderForStock(Sheet sheet, int startExcelRow, CellStyle allBorderGrayAlignCenter){
		//------------------------------------- Accum Stock -----------------------------------------//
    Cell accumStock1H = ExcelUtils.getCell("AM", startExcelRow, sheet);
    accumStock1H.setCellStyle(allBorderGrayAlignCenter);
    accumStock1H.setCellValue("Accum stock");
    
    Cell accumStock2H = ExcelUtils.getNextColumn(accumStock1H);
    accumStock2H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStock3H = ExcelUtils.getNextColumn(accumStock2H);
    accumStock3H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStock4H = ExcelUtils.getNextColumn(accumStock3H);
    accumStock4H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStock5H = ExcelUtils.getNextColumn(accumStock4H);
    accumStock5H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStock6H = ExcelUtils.getNextColumn(accumStock5H);
    accumStock6H.setCellStyle(allBorderGrayAlignCenter);
    ExcelUtils.mergeColumn(sheet, accumStock1H, accumStock5H);
    
    //------------------------------------- Accum Stock Last month -----------------------------------------//
    Cell accumStockLastMonth1H = ExcelUtils.getNextColumn(accumStock6H);
    accumStockLastMonth1H.setCellStyle(allBorderGrayAlignCenter);
    accumStockLastMonth1H.setCellValue("Accum stock Last month");
    
    Cell accumStockLastMonth2H = ExcelUtils.getNextColumn(accumStockLastMonth1H);
    accumStockLastMonth2H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStockLastMonth3H = ExcelUtils.getNextColumn(accumStockLastMonth2H);
    accumStockLastMonth3H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStockLastMonth4H = ExcelUtils.getNextColumn(accumStockLastMonth3H);
    accumStockLastMonth4H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStockLastMonth5H = ExcelUtils.getNextColumn(accumStockLastMonth4H);
    accumStockLastMonth5H.setCellStyle(allBorderGrayAlignCenter);
    
    Cell accumStockLastMonth6H = ExcelUtils.getNextColumn(accumStockLastMonth5H);
    accumStockLastMonth6H.setCellStyle(allBorderGrayAlignCenter);
    ExcelUtils.mergeColumn(sheet, accumStockLastMonth1H, accumStockLastMonth6H);
    
    //------------------------------------- Accum Back Order Last month -----------------------------------------//
    Cell accumGAPStock1H = ExcelUtils.getNextColumn(accumStockLastMonth6H);
    accumGAPStock1H.setCellStyle(allBorderGrayAlignCenter);
    accumGAPStock1H.setCellValue("Accum GAP Stock");
    
    Cell accumGAPStock2H = ExcelUtils.getNextColumn(accumGAPStock1H);
    accumGAPStock2H.setCellStyle(allBorderGrayAlignCenter);
    ExcelUtils.mergeColumn(sheet, accumGAPStock1H, accumGAPStock2H);
    //------------------------------------------------------------------------------------------------------------//
    
    Cell dummyCell = ExcelUtils.getNextColumn(accumGAPStock2H);
		dummyCell.setCellStyle(allBorderGrayAlignCenter);
		dummyCell.setCellValue("use for new line");
		sheet.setColumnHidden(dummyCell.getColumnIndex(), true);
	}
	
	protected void createCommonHeaderForStock(Sheet sheet, Cell beginCell, CellStyle allBorderGrayAlignCenter, CellStyle allBorderYellowAlignCenter){
		Cell accumStock1 = ExcelUtils.getNextColumn(beginCell);
		accumStock1.setCellStyle(allBorderGrayAlignCenter);
		accumStock1.setCellValue("Stock end month");    
    
    Cell accumStock2 = ExcelUtils.getNextColumn(accumStock1);
    accumStock2.setCellStyle(allBorderGrayAlignCenter);
    accumStock2.setCellValue("New WS");
    
    Cell accumStock3 = ExcelUtils.getNextColumn(accumStock2);
    accumStock3.setCellStyle(allBorderGrayAlignCenter);
    accumStock3.setCellValue("Total");
    
    Cell accumStock4 = ExcelUtils.getNextColumn(accumStock3);
    accumStock4.setCellStyle(allBorderYellowAlignCenter);
    accumStock4.setCellValue("WS Cross sales In");
    
    Cell accumStock5 = ExcelUtils.getNextColumn(accumStock4);
    accumStock5.setCellStyle(allBorderYellowAlignCenter);
    accumStock5.setCellValue("WS Cross sales Out");
    
    Cell accumStock6 = ExcelUtils.getNextColumn(accumStock5);
    accumStock6.setCellStyle(allBorderYellowAlignCenter);
    accumStock6.setCellValue("Total");
    
    
    
    Cell accumStockLastMonth1 = ExcelUtils.getNextColumn(accumStock6);
    accumStockLastMonth1.setCellStyle(allBorderGrayAlignCenter);
    accumStockLastMonth1.setCellValue("Stock end month");   
    
    Cell accumStockLastMonth2 = ExcelUtils.getNextColumn(accumStockLastMonth1);
    accumStockLastMonth2.setCellStyle(allBorderGrayAlignCenter);
    accumStockLastMonth2.setCellValue("New WS");
    
    Cell accumStockLastMonth3 = ExcelUtils.getNextColumn(accumStockLastMonth2);
    accumStockLastMonth3.setCellStyle(allBorderGrayAlignCenter);
    accumStockLastMonth3.setCellValue("Total");
    
    Cell accumStockLastMonth4 = ExcelUtils.getNextColumn(accumStockLastMonth3);
    accumStockLastMonth4.setCellStyle(allBorderYellowAlignCenter);
    accumStockLastMonth4.setCellValue("WS Cross sales In");
    
    Cell accumStockLastMonth5 = ExcelUtils.getNextColumn(accumStockLastMonth4);
    accumStockLastMonth5.setCellStyle(allBorderYellowAlignCenter);
    accumStockLastMonth5.setCellValue("WS Cross sales Out");
    
    Cell accumStockLastMonth6 = ExcelUtils.getNextColumn(accumStockLastMonth5);
    accumStockLastMonth6.setCellStyle(allBorderYellowAlignCenter);
    accumStockLastMonth6.setCellValue("Total");
    
    Cell accumGAPBackOrder1 = ExcelUtils.getNextColumn(accumStockLastMonth6);
    accumGAPBackOrder1.setCellStyle(allBorderGrayAlignCenter);
    accumGAPBackOrder1.setCellValue("Previous month");   
    
    Cell accumGAPBackOrder2 = ExcelUtils.getNextColumn(accumGAPBackOrder1);
    accumGAPBackOrder2.setCellStyle(allBorderGrayAlignCenter);
    accumGAPBackOrder2.setCellValue("This month");	
	}
	
	
	
	public Cell createCommonHeader(Sheet sheet, Cell dayCol, ReportCriteria criteria){
    CellStyle allBorderYellowAlignCenter = StyleUtils.allBorderYellowAlignCenter(sheet.getWorkbook().createCellStyle());
    allBorderYellowAlignCenter.setFont(FontUtils.blackBold(sheet));
    
    CellStyle allBorderRedAlignCenter = StyleUtils.allBorderRedAlignCenter(sheet.getWorkbook().createCellStyle());
    allBorderYellowAlignCenter.setFont(FontUtils.blackBold(sheet));
    
    CellStyle allBorderBuleAlignCenter = StyleUtils.allBorderBlueAlignCenter(sheet.getWorkbook().createCellStyle());
    allBorderBuleAlignCenter.setFont(FontUtils.whiteBold(sheet));
    
    //Cell dayCol = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    int[] days = DateUtils.getNumberOfDays(criteria.getReportDate());
    for(int day:days){
    	if(days[0] != (day)){
    		dayCol = ExcelUtils.getNextColumn(dayCol);
    	}
    	dayCol.setCellStyle(allBorderYellowAlignCenter);
    	dayCol.setCellValue(day);
    }
    
    Cell totalCal = ExcelUtils.getNextColumn(dayCol);
    totalCal.setCellStyle(allBorderRedAlignCenter);
    totalCal.setCellValue("Total");
    
		Cell lastMonthCol = ExcelUtils.getNextColumn(totalCal);
		lastMonthCol.setCellStyle(allBorderBuleAlignCenter);
		lastMonthCol.setCellValue("Last month");
		
		Cell vsLastMthCol = ExcelUtils.getNextColumn(lastMonthCol);
		vsLastMthCol.setCellStyle(allBorderBuleAlignCenter);
		vsLastMthCol.setCellValue("vs last mth");
		
		return vsLastMthCol;
	}
	
	public Cell createCommonFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int beginDataRow) {
		CellStyle allBorderYellow = StyleUtils.allBorderYellow2(sheet.getWorkbook().createCellStyle());
		allBorderYellow.setFont(FontUtils.blackBold(sheet));
		
    Cell footerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    footerColD.setCellStyle(allBorderYellow);
    footerColD.setCellValue("Total");
    sheet.autoSizeColumn(footerColD.getColumnIndex());
    
    int endDataRow = startExcelRow-1;
    List<Cell> dayCols = createDayCols(criteria, footerColD);
    allBorderYellow.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));
  	for(Cell dayCol:dayCols){
  		dayCol.setCellStyle(allBorderYellow);
  		dayCol.setCellType(CellType.FORMULA);
			dayCol.setCellFormula(FormulasUtils.sum(dayCol, beginDataRow, endDataRow));
		}
  	
  	Cell totalCol = ExcelUtils.getNextColumn(dayCols.get(dayCols.size()-1));
  	totalCol.setCellStyle(allBorderYellow);
  	totalCol.setCellType(CellType.FORMULA);
  	totalCol.setCellFormula(FormulasUtils.sum(totalCol, beginDataRow, endDataRow));
  	
		Cell lastMonthCol = ExcelUtils.getNextColumn(totalCol);
  	lastMonthCol.setCellStyle(allBorderYellow);
  	lastMonthCol.setCellType(CellType.FORMULA);
  	lastMonthCol.setCellFormula(FormulasUtils.sum(lastMonthCol, beginDataRow, endDataRow));
  	
  	Cell vsLastMthCol = ExcelUtils.getNextColumn(lastMonthCol);
  	vsLastMthCol.setCellStyle(allBorderYellow);
  	vsLastMthCol.setCellType(CellType.FORMULA);
  	vsLastMthCol.setCellFormula(FormulasUtils.sum(vsLastMthCol, beginDataRow, endDataRow));
  	return vsLastMthCol;
	}
	
//---------------------------------------------------- For Back Order ----------------------------------------------------//	
	protected void createCommonFooterForBackOrder(Sheet sheet,Cell beginCell, int beginDataRow) {
		CellStyle allBorderYellow = StyleUtils.allBorderYellow2(sheet.getWorkbook().createCellStyle());
		allBorderYellow.setFont(FontUtils.blackBold(sheet));
		
		int endDataRow = beginCell.getRowIndex();
		
		Cell accumPreviousMonthCol = ExcelUtils.getNextColumn(beginCell);
		accumPreviousMonthCol.setCellStyle(allBorderYellow);
		accumPreviousMonthCol.setCellType(CellType.FORMULA);
		accumPreviousMonthCol.setCellFormula(FormulasUtils.sum(accumPreviousMonthCol, beginDataRow, endDataRow));
  	
  	Cell accumThisMonthCol = ExcelUtils.getNextColumn(accumPreviousMonthCol);
  	accumThisMonthCol.setCellStyle(allBorderYellow);
  	accumThisMonthCol.setCellType(CellType.FORMULA);
  	accumThisMonthCol.setCellFormula(FormulasUtils.sum(accumThisMonthCol, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthPreviousMonthCol = ExcelUtils.getNextColumn(accumThisMonthCol);
  	accumLastMonthPreviousMonthCol.setCellStyle(allBorderYellow);
  	accumLastMonthPreviousMonthCol.setCellType(CellType.FORMULA);
  	accumLastMonthPreviousMonthCol.setCellFormula(FormulasUtils.sum(accumLastMonthPreviousMonthCol, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthThisMonthCol = ExcelUtils.getNextColumn(accumLastMonthPreviousMonthCol);
  	accumLastMonthThisMonthCol.setCellStyle(allBorderYellow);
  	accumLastMonthThisMonthCol.setCellType(CellType.FORMULA);
  	accumLastMonthThisMonthCol.setCellFormula(FormulasUtils.sum(accumLastMonthThisMonthCol, beginDataRow, endDataRow));
  	
  	Cell accumGAPPreviousMonthCol = ExcelUtils.getNextColumn(accumLastMonthThisMonthCol);
  	accumGAPPreviousMonthCol.setCellStyle(allBorderYellow);
  	accumGAPPreviousMonthCol.setCellType(CellType.FORMULA);
  	accumGAPPreviousMonthCol.setCellFormula(FormulasUtils.sum(accumGAPPreviousMonthCol, beginDataRow, endDataRow));
  	
  	Cell accumGAPThisMonth = ExcelUtils.getNextColumn(accumGAPPreviousMonthCol);
  	accumGAPThisMonth.setCellStyle(allBorderYellow);
  	accumGAPThisMonth.setCellType(CellType.FORMULA);
  	accumGAPThisMonth.setCellFormula(FormulasUtils.sum(accumGAPThisMonth, beginDataRow, endDataRow));
	}
	
//---------------------------------------------------- For Stock ----------------------------------------------------//	
	protected void createCommonFooterForStock(Sheet sheet,Cell beginCell, int beginDataRow) {
		CellStyle allBorderYellow = StyleUtils.allBorderYellow2(sheet.getWorkbook().createCellStyle());
		allBorderYellow.setFont(FontUtils.blackBold(sheet));
		
		int endDataRow = beginCell.getRowIndex();
		
		Cell accumStockEndLastMthCol = ExcelUtils.getNextColumn(beginCell);
		accumStockEndLastMthCol.setCellStyle(allBorderYellow);
		accumStockEndLastMthCol.setCellType(CellType.FORMULA);
		accumStockEndLastMthCol.setCellFormula(FormulasUtils.sum(accumStockEndLastMthCol, beginDataRow, endDataRow));
  	
  	Cell accumNewWSCol = ExcelUtils.getNextColumn(accumStockEndLastMthCol);
  	accumNewWSCol.setCellStyle(allBorderYellow);
  	accumNewWSCol.setCellType(CellType.FORMULA);
  	accumNewWSCol.setCellFormula(FormulasUtils.sum(accumNewWSCol, beginDataRow, endDataRow));
  	
  	Cell accumTotal1Col = ExcelUtils.getNextColumn(accumNewWSCol);
  	accumTotal1Col.setCellStyle(allBorderYellow);
  	accumTotal1Col.setCellType(CellType.FORMULA);
  	accumTotal1Col.setCellFormula(FormulasUtils.sum(accumTotal1Col, beginDataRow, endDataRow));
  	
  	Cell accumWSCrossSaleInCol = ExcelUtils.getNextColumn(accumTotal1Col);
  	accumWSCrossSaleInCol.setCellStyle(allBorderYellow);
  	accumWSCrossSaleInCol.setCellType(CellType.FORMULA);
  	accumWSCrossSaleInCol.setCellFormula(FormulasUtils.sum(accumWSCrossSaleInCol, beginDataRow, endDataRow));
  	
  	Cell accumWSCrossSaleOutCol = ExcelUtils.getNextColumn(accumWSCrossSaleInCol);
  	accumWSCrossSaleOutCol.setCellStyle(allBorderYellow);
  	accumWSCrossSaleOutCol.setCellType(CellType.FORMULA);
  	accumWSCrossSaleOutCol.setCellFormula(FormulasUtils.sum(accumWSCrossSaleOutCol, beginDataRow, endDataRow));
  	
  	Cell accumTotal2Col = ExcelUtils.getNextColumn(accumWSCrossSaleOutCol);
  	accumTotal2Col.setCellStyle(allBorderYellow);
  	accumTotal2Col.setCellType(CellType.FORMULA);
  	accumTotal2Col.setCellFormula(FormulasUtils.sum(accumTotal2Col, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthStockEndLastMthCol = ExcelUtils.getNextColumn(accumTotal2Col);
  	accumLastMonthStockEndLastMthCol.setCellStyle(allBorderYellow);
  	accumLastMonthStockEndLastMthCol.setCellType(CellType.FORMULA);
  	accumLastMonthStockEndLastMthCol.setCellFormula(FormulasUtils.sum(accumLastMonthStockEndLastMthCol, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthNewWSCol = ExcelUtils.getNextColumn(accumLastMonthStockEndLastMthCol);
  	accumLastMonthNewWSCol.setCellStyle(allBorderYellow);
  	accumLastMonthNewWSCol.setCellType(CellType.FORMULA);
  	accumLastMonthNewWSCol.setCellFormula(FormulasUtils.sum(accumLastMonthNewWSCol, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthTotal1Col = ExcelUtils.getNextColumn(accumLastMonthNewWSCol);
  	accumLastMonthTotal1Col.setCellStyle(allBorderYellow);
  	accumLastMonthTotal1Col.setCellType(CellType.FORMULA);
  	accumLastMonthTotal1Col.setCellFormula(FormulasUtils.sum(accumLastMonthTotal1Col, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthWSCrossSaleInCol = ExcelUtils.getNextColumn(accumLastMonthTotal1Col);
  	accumLastMonthWSCrossSaleInCol.setCellStyle(allBorderYellow);
  	accumLastMonthWSCrossSaleInCol.setCellType(CellType.FORMULA);
  	accumLastMonthWSCrossSaleInCol.setCellFormula(FormulasUtils.sum(accumLastMonthWSCrossSaleInCol, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthWSCrossSaleOutCol = ExcelUtils.getNextColumn(accumLastMonthWSCrossSaleInCol);
  	accumLastMonthWSCrossSaleOutCol.setCellStyle(allBorderYellow);
  	accumLastMonthWSCrossSaleOutCol.setCellType(CellType.FORMULA);
  	accumLastMonthWSCrossSaleOutCol.setCellFormula(FormulasUtils.sum(accumLastMonthWSCrossSaleOutCol, beginDataRow, endDataRow));
  	
  	Cell accumLastMonthTotal2Col = ExcelUtils.getNextColumn(accumLastMonthWSCrossSaleOutCol);
  	accumLastMonthTotal2Col.setCellStyle(allBorderYellow);
  	accumLastMonthTotal2Col.setCellType(CellType.FORMULA);
  	accumLastMonthTotal2Col.setCellFormula(FormulasUtils.sum(accumLastMonthTotal2Col, beginDataRow, endDataRow));
  	
  	
  	Cell accumGAPPreviousMonthCol = ExcelUtils.getNextColumn(accumLastMonthTotal2Col);
  	accumGAPPreviousMonthCol.setCellStyle(allBorderYellow);
  	accumGAPPreviousMonthCol.setCellType(CellType.FORMULA);
  	accumGAPPreviousMonthCol.setCellFormula(FormulasUtils.sum(accumGAPPreviousMonthCol, beginDataRow, endDataRow));
  	
  	Cell accumGAPThisMonth = ExcelUtils.getNextColumn(accumGAPPreviousMonthCol);
  	accumGAPThisMonth.setCellStyle(allBorderYellow);
  	accumGAPThisMonth.setCellType(CellType.FORMULA);
  	accumGAPThisMonth.setCellFormula(FormulasUtils.sum(accumGAPThisMonth, beginDataRow, endDataRow));
	}
	
	
	protected List<Cell> createDayCols(ReportCriteria criteria,Cell startCol){
		List<Cell> dayCols = new ArrayList<Cell>();
		int[] days = DateUtils.getNumberOfDays(criteria.getReportDate());
		for(int day:days){
			if(days[0] == (day)){
				dayCols.add(ExcelUtils.getNextColumn(startCol));
			}else{
				dayCols.add(ExcelUtils.getNextColumn(dayCols.get(day-2)));
			}
		}
		return dayCols;
	}
	
	protected void createDayColsNextRow(List<Cell> dayCols){
		List<Cell> dayColsTmp = new ArrayList<Cell>();
		for(Cell dayCol:dayCols){
			dayCol= ExcelUtils.getNextRow(dayCol);
			dayColsTmp.add(dayCol);
		}
		dayCols.removeAll(dayCols);
		dayCols.addAll(dayColsTmp);
		dayColsTmp.removeAll(dayColsTmp);
	}
	
	protected void setDayColsValue(List<Cell> dayCols, CellStyle allBorderAlignRight, Object stockModel){
		int dayNum = 1;
		String colValue;
  	for(Cell dayCol:dayCols){
  		colValue = String.valueOf(Utils.getValueByMethodName(stockModel, "day"+dayNum));
			dayCol.setCellStyle(allBorderAlignRight);
			dayCol.setCellValue(Utils.convertString2Int(colValue));
  		dayNum++;
		}
	}
	
	protected String getDiffReturn0WhenMinus(Cell firstCell, Cell nextCells){
		CellReference firstCellRef = new CellReference(firstCell);
		CellReference nextCellRef = new CellReference(nextCells);
		String formulas = firstCellRef.formatAsString()+"-"+nextCellRef.formatAsString();
		return "IF("+formulas+"<0,0,"+formulas+")";
	}
}
