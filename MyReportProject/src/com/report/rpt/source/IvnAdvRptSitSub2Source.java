package com.report.rpt.source;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.InvAdvRptSitDAO;
import com.report.model.ChannelModel;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class IvnAdvRptSitSub2Source extends IvnAdvRptSitCommonSource implements ReportSource {
	
	@Override
	public InvAdvRptSitDAO getDAOobject() {
		return (InvAdvRptSitDAO)Utils.APP_CONTEXT.getBean("invAdvRptSitDAO");
	}
	
	@Override
	public void createHeader(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria){
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		Cell headerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		headerColD.setCellStyle(allBorder);
		headerColD.setCellValue("Area");
	}
	
	@Override
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria){
    List<ChannelModel> channelModels = getDAOobject().getSubReport2Data();
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    Cell bodyColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    for(ChannelModel channelModel:channelModels){
    	
    	bodyColD.setCellStyle(allBorder);
    	bodyColD.setCellValue(channelModel.getAreaId());
    	ExcelUtils.createDummyColumn(bodyColD, allBorder, 36);
    	
    	if(channelModel != channelModels.get(channelModels.size()-1)){
    		bodyColD = ExcelUtils.getNextRow(bodyColD);
    	}
    }
    
	}
	
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int... params){
		CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
		Cell footerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		footerColD.setCellStyle(allBorderYellow);
		footerColD.setCellValue("Total");
		ExcelUtils.createDummyColumn(footerColD, allBorderYellow, 36);
	}
	
}
