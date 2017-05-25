package com.report.rpt.source;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import com.report.dao.InvAdvRptSitDAO;
import com.report.model.MastCarSeriesModel;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.ExcelUtils;
import com.report.util.FormulasUtils;
import com.report.util.StyleUtils;
import com.report.util.Utils;

public class IvnAdvRptSitSub1Source extends IvnAdvRptSitCommonSource implements ReportSource {

	@Override
	public InvAdvRptSitDAO getDAOobject() {
		return (InvAdvRptSitDAO)Utils.APP_CONTEXT.getBean("invAdvRptSitDAO");
	}
	
	@Override
	public void createHeader(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria){
		CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
		Cell dummyHeaderColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
		dummyHeaderColD.setCellStyle(allBorder);
		
    Cell headerColD = ExcelUtils.getCell(startExcelColumn, (startExcelRow+1), sheet);
    headerColD.setCellStyle(allBorder);
    headerColD.setCellValue("Model");
	}
	
	@Override
	public void createBody(Sheet sheet, String startExcelColumn,int startExcelRow, ReportCriteria criteria){
    List<MastCarSeriesModel> mastCarSeriesModels = getDAOobject().getSubReport1Data();
    CellStyle allBorder = StyleUtils.allBorder(sheet.getWorkbook().createCellStyle());
    
    Cell bodyColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    for(MastCarSeriesModel mastCarSeriesModel:mastCarSeriesModels){
  		
    	bodyColD.setCellStyle(allBorder);
  		bodyColD.setCellValue(mastCarSeriesModel.getSeriesEname());
    	ExcelUtils.createDummyColumn(bodyColD, allBorder, 36);
    	
    	if(mastCarSeriesModel != mastCarSeriesModels.get(mastCarSeriesModels.size()-1)){
    		bodyColD = ExcelUtils.getNextRow(bodyColD);
    	}
    }
	}
	
	@Override
	public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int... params) {
		CellStyle allBorderYellow = StyleUtils.allBorderYellow(sheet.getWorkbook().createCellStyle());
		
    Cell footerColD = ExcelUtils.getCell(startExcelColumn, startExcelRow, sheet);
    footerColD.setCellStyle(allBorderYellow);
    footerColD.setCellValue("Total");
    
    Cell footerColE = ExcelUtils.getNextColumn(footerColD);
    footerColE.setCellStyle(allBorderYellow);
    footerColE.setCellType(CellType.FORMULA);
    //footerColE.setCellFormula("SUM(E"+params[0]+":E"+(startExcelRow-1)+")");
    footerColE.setCellFormula(FormulasUtils.sum(footerColE, params[0], startExcelRow-1));
    ExcelUtils.createDummyColumn(footerColE, allBorderYellow, 35);
		//ExcelUtils.createDummyColumn(footerColD, allBorderYellow, 36);
	}

	
}
