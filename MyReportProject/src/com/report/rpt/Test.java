package com.report.rpt;

import org.apache.poi.ss.usermodel.Workbook;

import com.report.generator.ReportGenerator;
import com.report.rpt.criteria.ReportCriteria;

public class Test implements ReportGenerator{

	@Override
	public String getFilePath(ReportCriteria reportCriteria) {
		return null;
	}

	@Override
	public Workbook getWorkbook(ReportCriteria reportCriteria) {
		return null;
	}


}
