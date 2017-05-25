package com.report.rpt.criteria;

import java.util.Date;

import com.report.util.DateUtils;

public class ReportCriteria {
	
	public Date getReportDate() {
		return DateUtils.getYesterday();
	}

}
