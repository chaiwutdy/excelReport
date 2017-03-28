package com.report.constant;

import com.report.generator.ReportGenerator;
import com.report.rpt.DummyRpt;
import com.report.rpt.IvnAdvRptSit;

public enum ReportType {
	DUMMY_RPT (0,new DummyRpt()),
	IVN_ADV_RPT_SIT (1, new IvnAdvRptSit());
	
	private final int id;
	private final ReportGenerator reportGenerator;
	
	ReportType(int id, ReportGenerator reportGenerator){
		this.id = id;
		this.reportGenerator = reportGenerator;
	}
	
	public int getId() {
		return id;
	}

	public ReportGenerator getReportGenerator() {
		return reportGenerator;
	}

}
