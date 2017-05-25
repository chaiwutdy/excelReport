package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Sheet;

import com.report.rpt.criteria.ReportCriteria;

/**
 * This Interface to define The structure of a ReportSource.<br>
 * All ReportSource need to implements this interface.
 */
public interface ReportSource {
	
	/**
	 * This method always return Object.<br>
	 * To use it have to change return Object to DAO Object in dao Package.
	 * @return	DAO Object in com.report.dao Package
	 */
	public Object getDAOobject();
	
	/**
	 * This method for creating header.
	 * @param		sheet	This is the first parameter to use create header parts. 
	 * @param		startExcelColumn	This is the second parameter to specify the excel column where data starts.<br>
	 * 					Exp. "A" or "B" or "C"
	 * @param		startExcelRow	This is the third parameter to specify the excel row where data starts.<br>
	 * 					Exp. 1 or 2 or 3
	 */
	public void createHeader(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria);
	
	/**
	 * This method for creating body.
	 * @param		sheet	This is the first parameter to use create body parts.
	 * @param		startExcelColumn	This is the second parameter to specify the excel column where data starts.<br>
	 * 				  Exp. "A" or "B" or "C"
	 * @param		startExcelRow	This is the third parameter to specify the excel row where data starts.<br>
	 * 					Exp. 1 or 2 or 3
	 */
	public void createBody(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria);
	
	/**
	 * This method for creating footer.<br>
	 * This method has optional integer parameters<br>
	 * How to call this method ?<br>
	 * Exp. object.createFooter(sheet, startExcelColumn, startExcelRow); or <br>
	 * 			object.createFooter(sheet, startExcelColumn, startExcelRow, startBodyExcelRow); or <br>
	 * 			object.createFooter(sheet, startExcelColumn, startExcelRow, startHeadExcelRow, startBodyExcelRow); <br>
	 * How to get Optional Parameters ?<br>
	 * Exp. public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, int startBodyExcelRow){<br>
	 * 					System.out.println("startBodyExcelRow:"startBodyExcelRow[0]); //Getting value <br>
	 * 			}<br>
	 * @param		sheet	This is the first parameter to use create footer parts.
	 * @param		startExcelColumn	This is the second parameter to specify the excel column where data starts.<br>
	 *  				Exp. "A" or "B" or "C"
	 * @param		startExcelRow	This is the third parameter to specify the excel row where data starts.<br>
	 * 					Exp. 1 or 2 or 3
	 * @param		params	This is Optional Parameters.<br>
	 * 					To get this parameters value like get value from array.<br>				
	 */
	public void createFooter(Sheet sheet, String startExcelColumn, int startExcelRow, ReportCriteria criteria, int... params);
}
