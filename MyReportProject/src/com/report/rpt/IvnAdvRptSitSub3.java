package com.report.rpt;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.model.ChannelModel;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;

public class IvnAdvRptSitSub3 extends IvnAdvRptSitMain {
	
	public void createHeader(Sheet sheet){
		int startRowNum = ExcelUtils.getLastRow(sheet)+2;
		int headerRowNum = startRowNum+1;
		String commonHeaderCellRef = "E"+startRowNum;
		
		createCommonHeader(sheet,commonHeaderCellRef);
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
	
	public void createBody(Sheet sheet){
    List<ChannelModel> channelModels = getDAOobject().getSubReport3Data();
	    
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
    
    String startBodyCellRef ="A"+(ExcelUtils.getLastRow(sheet)+1);
    Cell bodyColA = null;
    Cell bodyColB = null;
    Cell bodyColC = null;
    Cell bodyColD = null;
    Cell dummyCell = null;
    for(ChannelModel channelModel:channelModels){
    	if(channelModels.get(0) == channelModel){
    		bodyColA = ExcelUtils.getCell(startBodyCellRef, sheet);
    		bodyColA.setCellStyle(allBorder);
    		bodyColA.setCellValue(channelModel.getZoneId());
    		
    		bodyColB = ExcelUtils.getNextColumn(bodyColA, sheet);
    		bodyColB.setCellStyle(allBorder);
    		bodyColB.setCellValue(channelModel.getAreaId());
    		
    		bodyColC = ExcelUtils.getNextColumn(bodyColB, sheet);
    		bodyColC.setCellStyle(allBorder);
    		bodyColC.setCellValue(channelModel.getChannelCode());
    		
    		bodyColD = ExcelUtils.getNextColumn(bodyColC, sheet);
    		bodyColD.setCellStyle(allBorder);
    		bodyColD.setCellValue(channelModel.getChShortName());
    		
    	}else{
    		bodyColA = ExcelUtils.getNextRow(bodyColA, sheet);
    		bodyColA.setCellStyle(allBorder);
    		bodyColA.setCellValue(channelModel.getZoneId());
    		
    		bodyColB = ExcelUtils.getNextRow(bodyColB, sheet);
    		bodyColB.setCellStyle(allBorder);
    		bodyColB.setCellValue(channelModel.getAreaId());
    		
    		bodyColC = ExcelUtils.getNextRow(bodyColC, sheet);
    		bodyColC.setCellStyle(allBorder);
    		bodyColC.setCellValue(channelModel.getChannelCode());
    		
    		bodyColD = ExcelUtils.getNextRow(bodyColD, sheet);
    		bodyColD.setCellStyle(allBorder);
    		bodyColD.setCellValue(channelModel.getChShortName());
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
	    
    bodyColA = ExcelUtils.getNextRow(bodyColA, sheet);
    bodyColA.setCellStyle(allBorder);
    
    bodyColB = ExcelUtils.getNextRow(bodyColB, sheet);
    bodyColB.setCellStyle(allBorder);
    
    bodyColC = ExcelUtils.getNextRow(bodyColC, sheet);
    bodyColC.setCellStyle(allBorderYellow);
    bodyColC.setCellValue("Total");
		
		//generate dummy
		for(int i=0;i<=35;i++){
			bodyColC = ExcelUtils.getNextColumn(bodyColC, sheet);
			bodyColC.setCellStyle(allBorderYellow);
		}
	}
	
}
