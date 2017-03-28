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

public class IvnAdvRptSitSourceSub2 implements ReportSource {
	
	@Override
	public InvAdvRptSitDAO getDAOobject() {
		return (InvAdvRptSitDAO)Utils.APP_CONTEXT.getBean("invAdvRptSitDAO");
	}
	
	@Override
	public void createHeader(Sheet sheet, int startRowNum){
		String headerCellRef = "D"+(startRowNum+1);
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		Cell headerColD = ExcelUtils.getCell(headerCellRef, sheet);
		headerColD.setCellStyle(allBorder);
		headerColD.setCellValue("Area");
	}
	
	@Override
	public void createBody(Sheet sheet, int startRowNum){
    List<ChannelModel> channelModels = getDAOobject().getSubReport2Data();
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    
    String startBodyCellRef ="D"+startRowNum;
    Cell bodyColD = ExcelUtils.getCell(startBodyCellRef, sheet);
    for(ChannelModel channelModel:channelModels){
    	
    	if(channelModels.get(0) != channelModel){
    		bodyColD = ExcelUtils.getNextRow(bodyColD);
    	}
    	
    	bodyColD.setCellStyle(allBorder);
    	bodyColD.setCellValue(channelModel.getAreaId());
    	ExcelUtils.createDummyColumn(bodyColD, allBorder, 36);
    }
    
	}
	
	@Override
	public void createFooter(Sheet sheet, int startRowNum){
		CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
		
		String startFooterCellRef ="D"+startRowNum;
		Cell footerColD = ExcelUtils.getCell(startFooterCellRef, sheet);
		footerColD.setCellStyle(allBorderYellow);
		footerColD.setCellValue("Total");
		ExcelUtils.createDummyColumn(footerColD, allBorderYellow, 36);
	}
	
}
