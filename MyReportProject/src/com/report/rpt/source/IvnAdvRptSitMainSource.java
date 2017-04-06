package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.util.DateUtils;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;

public class IvnAdvRptSitMainSource implements ReportSource{

	@Override
	public Object getDAOobject() {
		return null;
	}

	@Override
	public void createHeader(Sheet sheet, String startExcelColumn,int startExcelRow) {
		Cell reportDate = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    reportDate.setCellValue("INVOICE ADVANCE SITUATION AS "+DateUtils.getCurrentDate("dd-MM-yyyy"));
	}

	@Override
	public void createBody(Sheet sheet, String startExcelColumn,int startExcelRow) {
	}
	
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn,int startExcelRow, int... params) {
	}
	
	public void createCommonHeader(Sheet sheet, String startExcelColumn, int startExcelRow){
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    CellStyle allBorderAlignCenter = StyleUtils.allBorderAlignCenter(sheet.getWorkbook().createCellStyle());
    
    Cell monthCol1 = null;
    Cell monthCol2 = null;
    Cell monthCol3 = null;
    Cell headerCell = null;
    String[] months = DateUtils.getShortMonths();
    String[] headers = {"Invoice","Delivery","Remain"};
    monthCol1 = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    for(String month:months){
    	
    	if(!months[0].equalsIgnoreCase(month)){
    		monthCol1 = ExcelUtils.getNextColumn(monthCol3);
    	}
    	monthCol1.setCellStyle(allBorderAlignCenter);
  		monthCol1.setCellValue(month);
    	
    	monthCol2 = ExcelUtils.getNextColumn(monthCol1);
    	monthCol2.setCellStyle(allBorder);
      
    	monthCol3 = ExcelUtils.getNextColumn(monthCol2);
    	monthCol3.setCellStyle(allBorder);
    	
      ExcelUtils.mergeColumn(sheet, monthCol1, monthCol3);
      
      for(String header:headers){
      	if(headers[0].equalsIgnoreCase(header) && months[0].equalsIgnoreCase(month)){
        	headerCell = ExcelUtils.getNextRow(monthCol1);//Utils.getCell(startHeadersCell, sheet);
      	}else{
      		headerCell = ExcelUtils.getNextColumn(headerCell);
      	}
      	
      	headerCell.setCellStyle(allBorderAlignCenter);
      	headerCell.setCellValue(header);
      }
      
    }
	}
}
