package com.report.generator;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;

import com.report.constant.ReportType;
import com.report.rpt.criteria.ReportCriteria;

/**
 * This class to generate reports that implements ReportGenerator class.
 */
public class GenerateReport {
	private static final Logger LOGGER = LogManager.getLogger(GenerateReport.class.getName());

	/**
	 * This method to write report that implements ReportGenerator class.
	 * @param		pReportType	This is the parameter to specify type of report.
	 * @return	a path of report file
	 */
	public String writeReport(ReportType pReportType){
		String filePath = null;
		Workbook book = null;
		try {
			ReportGenerator reportGenerator = getReport(pReportType);
			ReportCriteria reportCriteria = getReportCriteria(pReportType);
			if(reportGenerator!=null){
				filePath = reportGenerator.getFilePath(reportCriteria);
				
				if(filePath != null){
					book = reportGenerator.getWorkbook(reportCriteria);
					book.write(new FileOutputStream(filePath));
				}else{
					LOGGER.info("Please specify filePath");
				}
				
			}else{
				LOGGER.info("ReportType Not Found");
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally{ 
			try {
				if(book!=null){
					book.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return filePath;
	}
	
	
	/**
	 * This method to find ReportGenerator by ReportType.<br>
	 * @param		pReportType	This is the parameter to find ReportGenerator.
	 * @return	Object that implements ReportGenerator
	 */
	private ReportGenerator getReport(ReportType pReportType){
		for(ReportType reportType:ReportType.values()){
			if(reportType.getId() == pReportType.getId()){
				return reportType.getReportGenerator();
			}
		}
		return null;
	}
	
	private ReportCriteria getReportCriteria(ReportType pReportType){
		for(ReportType reportType:ReportType.values()){
			if(reportType.getId() == pReportType.getId()){
				return reportType.getReportCriteria();
			}
		}
		return null;
	}
	
}
