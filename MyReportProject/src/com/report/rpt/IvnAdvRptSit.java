package com.report.rpt;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.generator.ReportGenerator;
import com.report.rpt.criteria.ReportCriteria;
import com.report.rpt.source.IvnAdvRptSitMainSource;
import com.report.rpt.source.IvnAdvRptSitSub1Source;
import com.report.rpt.source.IvnAdvRptSitSub2Source;
import com.report.rpt.source.IvnAdvRptSitSub3Source;
import com.report.util.ExcelUtils;
import com.report.util.PixelUtils;
import com.report.util.Utils;

public class IvnAdvRptSit implements ReportGenerator{

	@Override
	public String getFilePath(ReportCriteria reportCriteria) {
		return Utils.getProperties("app.path.invAdvRptsit");
	}

	@Override
	public Workbook getWorkbook(ReportCriteria reportCriteria) {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("Sheet1");
    sheet.setColumnWidth(ExcelUtils.getCell("D1",sheet).getColumnIndex(), PixelUtils.widthUnits2Pixel((float)18.3));
    
		IvnAdvRptSitMainSource main = new IvnAdvRptSitMainSource();
		main.createHeader(sheet,"A",2,reportCriteria);
		//main.createBody(sheet);
		
		int startSub1RowNum = ExcelUtils.getLastRow(sheet)+4;
		IvnAdvRptSitSub1Source sub1 = new IvnAdvRptSitSub1Source();
		sub1.createCommonHeader(sheet, "E", startSub1RowNum);
		sub1.createHeader(sheet, "D", startSub1RowNum,null);
		int startBodyRowNum = ExcelUtils.getLastRow(sheet)+1;
		sub1.createBody(sheet, "D", startBodyRowNum,null);
		sub1.createFooter(sheet, "D", ExcelUtils.getLastRow(sheet)+1,null,startBodyRowNum);
		
		int startSub2RowNum = ExcelUtils.getLastRow(sheet)+4;
		IvnAdvRptSitSub2Source sub2 = new IvnAdvRptSitSub2Source();
		sub2.createCommonHeader(sheet, "E", startSub2RowNum);
		sub2.createHeader(sheet, "D", startSub2RowNum+1, null);
		sub2.createBody(sheet, "D", ExcelUtils.getLastRow(sheet)+1, null);
		sub2.createFooter(sheet, "D", ExcelUtils.getLastRow(sheet)+1, null);
		
		int startSub3RowNum = ExcelUtils.getLastRow(sheet)+2;
		IvnAdvRptSitSub3Source sub3 = new IvnAdvRptSitSub3Source();
		sub3.createCommonHeader(sheet, "E", startSub3RowNum);
		sub3.createHeader(sheet, "A", startSub3RowNum+1, null);
		sub3.createBody(sheet, "A", ExcelUtils.getLastRow(sheet)+1, null);
		sub3.createFooter(sheet, "A", ExcelUtils.getLastRow(sheet)+1, null);
		
		//main.createFooter(sheet);
		return book;
	}

}
