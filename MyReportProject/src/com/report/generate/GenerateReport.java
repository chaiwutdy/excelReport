package com.report.generate;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

import com.report.constant.ReportType;
import com.report.rpt.ReportBase;


public class GenerateReport {

	public String writeReport(ReportType pReportType){
		String filePath = null;
		Workbook book = null;
		try {
			ReportBase report = getReport(pReportType);
			if(report!=null){
				filePath = report.getFilePath();
				book = report.getWorkbook();
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
	
	public ReportBase getReport(ReportType pReportType){
		for(ReportType reportType:ReportType.values()){
			if(reportType.getId() == pReportType.getId()){
				return reportType.getReportBase();
			}
		}
		return null;
	}
	
}
