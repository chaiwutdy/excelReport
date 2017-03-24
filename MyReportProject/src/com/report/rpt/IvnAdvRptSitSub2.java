package com.report.rpt;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.model.ChannelModel;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;

public class IvnAdvRptSitSub2 extends IvnAdvRptSitMain {
	
	public void createHeader(Sheet sheet){
		int startRowNum = ExcelUtils.getLastRow(sheet)+4;
		String commonHeaderCellRef = "E"+startRowNum;
		String headerCellRef = "D"+(startRowNum+1);
		
		createCommonHeader(sheet, commonHeaderCellRef);
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		Cell headerColD = ExcelUtils.getCell(headerCellRef, sheet);
		headerColD.setCellStyle(allBorder);
		headerColD.setCellValue("Area");
	}
	
	public void createBody(Sheet sheet){
    List<ChannelModel> channelModels = getDAOobject().getSubReport2Data();
	    
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
    
    String startBodyCellRef ="D"+(ExcelUtils.getLastRow(sheet)+1);
    Cell bodyColD = null;
    Cell dummyCell = null;
    for(ChannelModel channelModel:channelModels){
    	if(channelModels.get(0) == channelModel){
    		bodyColD = ExcelUtils.getCell(startBodyCellRef, sheet);
    		bodyColD.setCellStyle(allBorder);
    		bodyColD.setCellValue(channelModel.getAreaId());
    		
    	}else{
    		bodyColD = ExcelUtils.getNextRow(bodyColD, sheet);
    		bodyColD.setCellStyle(allBorder);
    		bodyColD.setCellValue(channelModel.getAreaId());
    	}
    	
    	//generate dummy
    	for(int i=0;i<=35;i++){
    		if(i==0){
        		dummyCell = ExcelUtils.getNextColumn(bodyColD, sheet);
        		dummyCell.setCellStyle(allBorder);
    		}else{
    			dummyCell = ExcelUtils.getNextColumn(dummyCell, sheet);
        		dummyCell.setCellStyle(allBorder);
    		}
    	}
    }
    
    bodyColD = ExcelUtils.getNextRow(bodyColD, sheet);
		bodyColD.setCellStyle(allBorderYellow);
		bodyColD.setCellValue("Total");
		
		//generate dummy
		for(int i=0;i<=35;i++){
			bodyColD = ExcelUtils.getNextColumn(bodyColD, sheet);
			bodyColD.setCellStyle(allBorderYellow);
		}
	}
	
}
