package com.report.rpt;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.model.MastCarSeriesModel;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;

public class IvnAdvRptSitSub1 extends IvnAdvRptSitMain {

	public void createHeader(Sheet sheet){
		int startRowNum = ExcelUtils.getLastRow(sheet)+4;
		String headerCellRef1 = "D"+startRowNum;
		String commonHeaderCellRef = "E"+startRowNum;
		String headerCellRef2 = "D"+(startRowNum+1);
		
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		Cell dummyHeaderColD = ExcelUtils.getCell(headerCellRef1, sheet);
		dummyHeaderColD.setCellStyle(allBorder);
		
		createCommonHeader(sheet, commonHeaderCellRef);
    Cell headerColD = ExcelUtils.getCell(headerCellRef2, sheet);
    headerColD.setCellStyle(allBorder);
    headerColD.setCellValue("Model");
	}
	
	public void createBody(Sheet sheet){
    List<MastCarSeriesModel> mastCarSeriesModels = getDAOobject().getSubReport1Data();
    
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
    
    String startBodyCellRef ="D"+(ExcelUtils.getLastRow(sheet)+1);
    Cell bodyCalD = null;
    Cell dummyCell = null;
    for(MastCarSeriesModel mastCarSeriesModel:mastCarSeriesModels){
    	if(mastCarSeriesModels.get(0) == mastCarSeriesModel){
    		bodyCalD = ExcelUtils.getCell(startBodyCellRef, sheet);
    		bodyCalD.setCellStyle(allBorder);
    		bodyCalD.setCellValue(mastCarSeriesModel.getSeriesEname());
    		
    	}else{
    		bodyCalD = ExcelUtils.getNextRow(bodyCalD, sheet);
    		bodyCalD.setCellStyle(allBorder);
    		bodyCalD.setCellValue(mastCarSeriesModel.getSeriesEname());
    	}
    	
    	//generate dummy
    	for(int i=0;i<=35;i++){
    		if(i==0){
        		dummyCell = ExcelUtils.getNextColumn(bodyCalD, sheet);
        		dummyCell.setCellStyle(allBorder);
    		}else{
    			dummyCell = ExcelUtils.getNextColumn(dummyCell, sheet);
        		dummyCell.setCellStyle(allBorder);
    		}
    	}
    }
    
    bodyCalD = ExcelUtils.getNextRow(bodyCalD, sheet);
		bodyCalD.setCellStyle(allBorderYellow);
		bodyCalD.setCellValue("Total");
		
		//generate dummy
		for(int i=0;i<=35;i++){
			bodyCalD = ExcelUtils.getNextColumn(bodyCalD, sheet);
			bodyCalD.setCellStyle(allBorderYellow);
		}	
         
	}
	
}
