package com.report.constant;

import com.report.generator.ReportGenerator;
import com.report.mail.MailGenerator;
import com.report.mail.NDODailySalesMail;
import com.report.rpt.DummyRpt;
import com.report.rpt.IvnAdvRptSit;
import com.report.rpt.NDODailySalesRpt;
import com.report.rpt.criteria.NDODailySalesCriteria;
import com.report.rpt.criteria.ReportCriteria;

/**
 * This enum to define ReportType.<br>
 * How to define ReportType(EnumType)<br>
 * This is an example -> NAME_OF_REPORT ( runningNumber , Class that implements ReportGenerator )
 */
public enum ReportType {
	
	/*-------- Start Define ReportType(EnumType) --------*/
	DUMMY_RPT (0, new ReportCriteria(), new DummyRpt(), null),
	IVN_ADV_RPT_SIT (1, new ReportCriteria(),new IvnAdvRptSit(), null),
	NDO_DAILY_SALES_RPT (2, new NDODailySalesCriteria(), new NDODailySalesRpt(), new NDODailySalesMail());

	/*-------- End Define ReportType(EnumType) --------*/
	
	private final int id;
	private final ReportGenerator reportGenerator;
	private final MailGenerator mailGenerator;
	private final ReportCriteria reportCriteria;
	
	ReportType(int id, ReportCriteria reportCriteria,ReportGenerator reportGenerator, MailGenerator mailGenerator){
		this.id = id;
		this.reportCriteria = reportCriteria;
		this.reportGenerator = reportGenerator;
		this.mailGenerator = mailGenerator;
	}
	
	public int getId() {
		return id;
	}

	public ReportGenerator getReportGenerator() {
		return reportGenerator;
	}

	public MailGenerator getMailGenerator() {
		return mailGenerator;
	}

	public ReportCriteria getReportCriteria() {
		return reportCriteria;
	}
	
}
