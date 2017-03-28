package com.report.generator;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import com.report.constant.ReportType;


public class GenerateReport {

	public String writeReport(ReportType pReportType){
		String filePath = null;
		Workbook book = null;
		try {
			ReportGenerator reportGenerator = getReport(pReportType);
			if(reportGenerator!=null){
				filePath = reportGenerator.getFilePath();
				book = reportGenerator.getWorkbook();
				book.write(new FileOutputStream(filePath));
			}else{
				System.out.println("ReportType Not Found");
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
	
	public ReportGenerator getReport(ReportType pReportType){
		for(ReportType reportType:ReportType.values()){
			if(reportType.getId() == pReportType.getId()){
				return reportType.getReportGenerator();
			}
		}
		return null;
	}
	
}
