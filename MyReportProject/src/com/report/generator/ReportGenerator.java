package com.report.generator;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * This Interface to define The structure of a ReportGenerator.<br>
 * All ReportGenerator need to implements this interface.<br>
 * *** When implements this interface need to add it in com.iars.constant.ReportType.
 */
public interface ReportGenerator {
	
	/**{@inheritDoc}
	 * This method always return path of report file.<br>
	 * Path of report file should store on config.properties file<br>
	 * and get value by Utils.getProperties(variable in config.properties).
	 * @return	path of report file
	 */
	public String getFilePath();
	
	/**
	 * This method always return Workbook.<br>
	 * Workbook is contents to generate excel file.
	 * @return	Workbook of Report
	 */
	public Workbook getWorkbook();
}
