package com.report.rpt;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.dao.InvAdvRptSitDAO;
import com.report.util.DateUtils;
import com.report.util.ExcelUtils;
import com.report.util.PixelUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class IvnAdvRptSitMain implements ReportBase{

	@Override
	public String getFilePath() {
		return Utils.getProperties("app.path.invAdvRptsit");
	}

	@Override
	public InvAdvRptSitDAO getDAOobject() {
		return (InvAdvRptSitDAO)Utils.APP_CONTEXT.getBean("invAdvRptSitDAO");
	}

	@Override
	public void createHeader(Sheet sheet) {
		Cell reportDate = ExcelUtils.getCell("A2",sheet);
    reportDate.setCellValue("INVOICE ADVANCE SITUATION AS "+DateUtils.getCurrentDate("dd-MM-yyyy"));
	}

	@Override
	public void createBody(Sheet sheet) {
		IvnAdvRptSitSub1 ivnAdvRptSitSub1= new IvnAdvRptSitSub1();
    ivnAdvRptSitSub1.createHeader(sheet);
    ivnAdvRptSitSub1.createBody(sheet);
    
    IvnAdvRptSitSub2 ivnAdvRptSitSub2 = new IvnAdvRptSitSub2();
    ivnAdvRptSitSub2.createHeader(sheet);
    ivnAdvRptSitSub2.createBody(sheet);
    
    IvnAdvRptSitSub3 ivnAdvRptSitSub3 = new IvnAdvRptSitSub3();
    ivnAdvRptSitSub3.createHeader(sheet);
    ivnAdvRptSitSub3.createBody(sheet);
	}

	@Override
	public Workbook getWorkbook() {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("Sheet1");
    sheet.setColumnWidth(ExcelUtils.getCell("D1",sheet).getColumnIndex(), PixelUtils.widthUnits2Pixel(16));
    createHeader(sheet);
    createBody(sheet);
		return book;
	}
	
	protected void createCommonHeader(Sheet sheet, String startMonthsCell){
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    CellStyle allBorderAlignCenter = StyleUtils.allBorderAlignCenter(sheet.getWorkbook().createCellStyle());
    
    Cell monthCol1 = null;
    Cell monthCol2 = null;
    Cell monthCol3 = null;
    Cell headerCell = null;
    String[] months = DateUtils.getShortMonths();
    String[] headers = {"Invoice","Delivery","Remain"};
    for(String month:months){
    	
    	if(months[0].equalsIgnoreCase(month)){
    		monthCol1 = ExcelUtils.getCell(startMonthsCell, sheet);
    		monthCol1.setCellStyle(allBorderAlignCenter);
    		monthCol1.setCellValue(month);
    	}else{
    		monthCol1 = ExcelUtils.getNextColumn(monthCol3, sheet);
    		monthCol1.setCellStyle(allBorderAlignCenter);
    		monthCol1.setCellValue(month);
    	}
    	
    	monthCol2 = ExcelUtils.getNextColumn(monthCol1, sheet);
    	monthCol2.setCellStyle(allBorder);
      
    	monthCol3 = ExcelUtils.getNextColumn(monthCol2,sheet);
    	monthCol3.setCellStyle(allBorder);
      ExcelUtils.mergeCell(sheet, monthCol1.getRowIndex(), monthCol1.getColumnIndex(), monthCol3.getColumnIndex());
      
      for(String header:headers){
      	if(headers[0].equalsIgnoreCase(header) && months[0].equalsIgnoreCase(month)){
        	headerCell = ExcelUtils.getNextRow(monthCol1, sheet);//Utils.getCell(startHeadersCell, sheet);
        	headerCell.setCellStyle(allBorderAlignCenter);
        	headerCell.setCellValue(header);
      	}else{
      		headerCell = ExcelUtils.getNextColumn(headerCell, sheet);
        	headerCell.setCellStyle(allBorderAlignCenter);
        	headerCell.setCellValue(header);
      	}
      }
      
    }
	}

}
