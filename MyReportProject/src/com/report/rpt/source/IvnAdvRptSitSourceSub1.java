package com.report.rpt.source;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.InvAdvRptSitDAO;
import com.report.model.MastCarSeriesModel;
import com.report.util.ExcelUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class IvnAdvRptSitSourceSub1 implements ReportSource {

	@Override
	public InvAdvRptSitDAO getDAOobject() {
		return (InvAdvRptSitDAO)Utils.APP_CONTEXT.getBean("invAdvRptSitDAO");
	}
	
	@Override
	public void createHeader(Sheet sheet, int startRowNum){
		String headerCellRef1 = "D"+startRowNum;
		String headerCellRef2 = "D"+(startRowNum+1);
		
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		Cell dummyHeaderColD = ExcelUtils.getCell(headerCellRef1, sheet);
		dummyHeaderColD.setCellStyle(allBorder);
		
    Cell headerColD = ExcelUtils.getCell(headerCellRef2, sheet);
    headerColD.setCellStyle(allBorder);
    headerColD.setCellValue("Model");
	}
	
	@Override
	public void createBody(Sheet sheet, int startRowNum){
    List<MastCarSeriesModel> mastCarSeriesModels = getDAOobject().getSubReport1Data();
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    
    String startBodyCellRef ="D"+startRowNum;
    Cell bodyColD = ExcelUtils.getCell(startBodyCellRef, sheet);
    for(MastCarSeriesModel mastCarSeriesModel:mastCarSeriesModels){
    	
    	if(mastCarSeriesModels.get(0) != mastCarSeriesModel){
    		bodyColD = ExcelUtils.getNextRow(bodyColD);
    	}
  		
    	bodyColD.setCellStyle(allBorder);
  		bodyColD.setCellValue(mastCarSeriesModel.getSeriesEname());
    	ExcelUtils.createDummyColumn(bodyColD, allBorder, 36);
    }
	}
	
	@Override
	public void createFooter(Sheet sheet, int startRowNum) {
		CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
		
		String startFooterCellRef ="D"+startRowNum;
    Cell footerColD = ExcelUtils.getCell(startFooterCellRef, sheet);
    footerColD.setCellStyle(allBorderYellow);
    footerColD.setCellValue("Total");
    /*
    Cell footerColE = ExcelUtils.getCell("E"+(ExcelUtils.getLastRow(sheet)), sheet);
    footerColE.setCellStyle(allBorderYellow);
    footerColE.setCellType(CellType.FORMULA);
    footerColE.setCellFormula("SUM(E1:E"+(ExcelUtils.getLastRow(sheet)-1)+")");
    ExcelUtils.createDummyColumn(footerColE, allBorderYellow, 35);
    */
		ExcelUtils.createDummyColumn(footerColD, allBorderYellow, 36);
	}

	
}
