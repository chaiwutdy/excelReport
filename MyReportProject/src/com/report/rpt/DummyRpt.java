package com.report.rpt;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.generator.ReportGenerator;
import com.report.rpt.criteria.ReportCriteria;
import com.report.rpt.source.DummyRptSource;
import com.report.util.ExcelUtils;
import com.report.util.Utils;

/**
 * This class is a ReportGenerator Example.
 */
public class DummyRpt implements ReportGenerator{

	@Override
	public String getFilePath(ReportCriteria reportCriteria) {
		return Utils.getProperties("app.path.dummyRpt");
	}

	@Override
	public Workbook getWorkbook(ReportCriteria reportCriteria) {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("SheetDummy");
		DummyRptSource dummyRptSource = new DummyRptSource();
		dummyRptSource.createHeader(sheet,"A",1,reportCriteria);
		dummyRptSource.createBody(sheet,"B",2,null);
		dummyRptSource.createFooter(sheet,"B",ExcelUtils.getLastRow(sheet)+1,null,2);
		return book;
	}

}
