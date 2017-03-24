package com.report.constant;

import com.report.rpt.DummyRpt;
import com.report.rpt.IvnAdvRptSitMain;
import com.report.rpt.ReportBase;

public enum ReportType {
	DUMMY_RPT (0,new DummyRpt()),
	IVN_ADV_RPT_SIT (1, new IvnAdvRptSitMain());
	
	private final int id;
	private final ReportBase reportBase;
	
	ReportType(int id, ReportBase reportBase){
		this.id = id;
		this.reportBase = reportBase;
	}
	
	public int getId() {
		return id;
	}

	public ReportBase getReportBase() {
		return reportBase;
	}
	
}
