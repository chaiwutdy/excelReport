package com.report.run;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.report.constant.ReportType;
import com.report.generator.GenerateReport;

public class RunGenerateReport {

	private static final Logger LOGGER = LogManager.getLogger(RunGenerateReport.class.getName());
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		LOGGER.debug("Debug Message Logged !!!");
    
//    LOGGER.error("Error Message Logged !!!", new NullPointerException("NullError"));
		genAllReport();
//		genSingleReport();
	}
	
	public static void genSingleReport(){
		GenerateReport report = new GenerateReport();
		//String result = report.writeReport(ReportType.DUMMY_RPT);
		String result = report.writeReport(ReportType.IVN_ADV_RPT_SIT);
		//System.out.println("Success! FilePath:"+result);
		LOGGER.info("Success! FilePath:"+result);
	}
	
	public static void genAllReport(){
		String result;
		GenerateReport generateReport = new GenerateReport();
		for(ReportType reportType :ReportType.values()){
			  result = generateReport.writeReport(reportType);
			  //System.out.println("Success! FilePath:"+result);
			  LOGGER.info("Success! FilePath:"+result);
		}
	}
}
