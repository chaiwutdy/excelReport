package com.report.rpt.source;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.InvAdvRptSitDAO;
import com.report.model.ChannelModel;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class IvnAdvRptSitSourceSub3 implements ReportSource {
	
	@Override
	public InvAdvRptSitDAO getDAOobject() {
		return (InvAdvRptSitDAO)Utils.APP_CONTEXT.getBean("invAdvRptSitDAO");
	}
	
	@Override
	public void createHeader(Sheet sheet, int startRowNum){
		int headerRowNum = startRowNum+1;
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    Cell headerColA = ExcelUtils.getCell("A"+headerRowNum, sheet);
    headerColA.setCellStyle(allBorder);
    headerColA.setCellValue("Zone ID");
    
    Cell headerColB = ExcelUtils.getCell("B"+headerRowNum, sheet);
    headerColB.setCellStyle(allBorder);
    headerColB.setCellValue("Area ID");
    
    Cell headerColC = ExcelUtils.getCell("C"+headerRowNum, sheet);
    headerColC.setCellStyle(allBorder);
    headerColC.setCellValue("Code");
    
    Cell headerColD = ExcelUtils.getCell("D"+headerRowNum, sheet);
    headerColD.setCellStyle(allBorder);
    headerColD.setCellValue("Name");
        
	}
	
	@Override
	public void createBody(Sheet sheet, int startRowNum){
    List<ChannelModel> channelModels = getDAOobject().getSubReport3Data();
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    
    String startBodyCellRef ="A"+startRowNum;
    Cell bodyColA = ExcelUtils.getCell(startBodyCellRef, sheet);
    Cell bodyColB = ExcelUtils.getNextColumn(bodyColA);
    Cell bodyColC = ExcelUtils.getNextColumn(bodyColB);
    Cell bodyColD = ExcelUtils.getNextColumn(bodyColC);
    for(ChannelModel channelModel:channelModels){
    	
    	if(channelModels.get(0) != channelModel){
    		bodyColA = ExcelUtils.getNextRow(bodyColA);
    		bodyColB = ExcelUtils.getNextRow(bodyColB);
    		bodyColC = ExcelUtils.getNextRow(bodyColC);
    		bodyColD = ExcelUtils.getNextRow(bodyColD);
    	}
    	
    	bodyColA.setCellStyle(allBorder);
  		bodyColA.setCellValue(channelModel.getZoneId());
  		
  		bodyColB.setCellStyle(allBorder);
  		bodyColB.setCellValue(channelModel.getAreaId());
  		
  		bodyColC.setCellStyle(allBorder);
  		bodyColC.setCellValue(channelModel.getChannelCode());
  		
  		bodyColD.setCellStyle(allBorder);
  		bodyColD.setCellValue(channelModel.getChShortName());
    	
    	ExcelUtils.createDummyColumn(bodyColD, allBorder, 36);
    }
    
	}
	
	@Override
	public void createFooter(Sheet sheet, int startRowNum){
		CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		
		String startFooterCellRef ="A"+startRowNum;
		Cell footerColA = ExcelUtils.getCell(startFooterCellRef, sheet);
		footerColA.setCellStyle(allBorder);
		
		Cell footerColB = ExcelUtils.getNextColumn(footerColA);
		footerColB.setCellStyle(allBorder);
		 
		Cell footerColC = ExcelUtils.getNextColumn(footerColB);
		footerColC.setCellStyle(allBorderYellow);
		footerColC.setCellValue("Total");
		ExcelUtils.createDummyColumn(footerColC, allBorderYellow, 36);
	}
	
}
