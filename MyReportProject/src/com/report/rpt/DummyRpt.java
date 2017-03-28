package com.report.rpt;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.generator.ReportGenerator;
import com.report.rpt.source.DummyRptSource;
import com.report.util.Utils;

public class DummyRpt implements ReportGenerator{

	@Override
	public String getFilePath() {
		return Utils.getProperties("app.path.dummyRpt");
	}

	@Override
	public Workbook getWorkbook() {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("SheetDummy");
		DummyRptSource dummyRptSource = new DummyRptSource();
		dummyRptSource.createHeader(sheet,1);
		dummyRptSource.createBody(sheet,2);
		dummyRptSource.createFooter(sheet,3);
		return book;
	}

}
