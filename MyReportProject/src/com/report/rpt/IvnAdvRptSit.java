package com.report.rpt;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.generator.ReportGenerator;
import com.report.rpt.source.IvnAdvRptSitSourceMain;
import com.report.rpt.source.IvnAdvRptSitSourceSub1;
import com.report.rpt.source.IvnAdvRptSitSourceSub2;
import com.report.rpt.source.IvnAdvRptSitSourceSub3;
import com.report.util.ExcelUtils;
import com.report.util.PixelUtils;
import com.report.util.Utils;

public class IvnAdvRptSit implements ReportGenerator{

	@Override
	public String getFilePath() {
		return Utils.getProperties("app.path.invAdvRptsit");
	}

	@Override
	public Workbook getWorkbook() {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("Sheet1");
    sheet.setColumnWidth(ExcelUtils.getCell("D1",sheet).getColumnIndex(), PixelUtils.widthUnits2Pixel(16));
    
		IvnAdvRptSitSourceMain main = new IvnAdvRptSitSourceMain();
		main.createHeader(sheet,2);
		//main.createBody(sheet);
		
		int startSub1RowNum = ExcelUtils.getLastRow(sheet)+4;
		main.createCommonHeader(sheet, "E"+startSub1RowNum);
		IvnAdvRptSitSourceSub1 sub1 = new IvnAdvRptSitSourceSub1();
		sub1.createHeader(sheet, startSub1RowNum);
		sub1.createBody(sheet, ExcelUtils.getLastRow(sheet)+1);
		sub1.createFooter(sheet, ExcelUtils.getLastRow(sheet)+1);
		
		int startSub2RowNum = ExcelUtils.getLastRow(sheet)+4;
		main.createCommonHeader(sheet, "E"+startSub2RowNum);
		IvnAdvRptSitSourceSub2 sub2 = new IvnAdvRptSitSourceSub2();
		sub2.createHeader(sheet, startSub2RowNum);
		sub2.createBody(sheet, ExcelUtils.getLastRow(sheet)+1);
		sub2.createFooter(sheet, ExcelUtils.getLastRow(sheet)+1);
		
		int startSub3RowNum = ExcelUtils.getLastRow(sheet)+2;
		main.createCommonHeader(sheet, "E"+startSub3RowNum);
		IvnAdvRptSitSourceSub3 sub3 = new IvnAdvRptSitSourceSub3();
		sub3.createHeader(sheet, startSub3RowNum);
		sub3.createBody(sheet, ExcelUtils.getLastRow(sheet)+1);
		sub3.createFooter(sheet, ExcelUtils.getLastRow(sheet)+1);
		
		//main.createFooter(sheet);
		return book;
	}

}
