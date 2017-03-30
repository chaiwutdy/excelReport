package com.report.constant;

import com.report.generator.ReportGenerator;
import com.report.rpt.DummyRpt;
import com.report.rpt.IvnAdvRptSit;

/**
 * This enum to define ReportType.<br>
 * How to define ReportType(EnumType)<br>
 * This is an example -> NAME_OF_REPORT ( runningNumber , Class that implements ReportGenerator )
 */
public enum ReportType {
	
	/*-------- Start Define ReportType(EnumType) --------*/
	DUMMY_RPT (0,new DummyRpt()),
	IVN_ADV_RPT_SIT (1, new IvnAdvRptSit());

	/*-------- End Define ReportType(EnumType) --------*/
	
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
