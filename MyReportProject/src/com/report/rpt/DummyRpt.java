package com.report.rpt;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.generator.ReportGenerator;
import com.report.rpt.source.DummyRptSource;
import com.report.util.ExcelUtils;
import com.report.util.Utils;

/**
 * This class is a ReportGenerator Example.
 */
public class DummyRpt implements ReportGenerator{

	@Override
	public String getFilePath() {
		return Utils.getProperties("app.path.dummyRpt");
	}

	@Override
	public Workbook getWorkbook() {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("SheetDummy");
		System.out.println(sheet.getLeftCol());
		DummyRptSource dummyRptSource = new DummyRptSource();
		dummyRptSource.createHeader(sheet,"A",1);
		dummyRptSource.createBody(sheet,"B",2);
		dummyRptSource.createFooter(sheet,"B",ExcelUtils.getLastRow(sheet)+1,2);
		System.out.println(sheet.getLeftCol());
		return book;
	}

}
