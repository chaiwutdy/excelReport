package com.report.rpt.source;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * This Interface to define The structure of a ReportSource.<br>
 * All ReportSource need to implements this interface.
 */
public interface ReportSource {
	
	/**
	 * This method always return Object.<br>
	 * To use it have to change return Object to DAO Object in dao Package.
	 * @return	DAO Object in com.iars.dao Package
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
	public void createHeader(Sheet sheet, String startExcelColumn,int startExcelRow);
	
	/**
	 * This method for creating body.
	 * @param		sheet	This is the first parameter to use create body parts.
	 * @param		startExcelColumn	This is the second parameter to specify the excel column where data starts.<br>
	 * 				  Exp. "A" or "B" or "C"
	 * @param		startExcelRow	This is the third parameter to specify the excel row where data starts.<br>
	 * 					Exp. 1 or 2 or 3
	 */
	public void createBody(Sheet sheet, String startExcelColumn,int startExcelRow);
	
	/**
	 * This method for creating footer.
	 * @param		sheet	This is the first parameter to use create footer parts.
	 * @param		startExcelColumn	This is the second parameter to specify the excel column where data starts.<br>
	 *  				Exp. "A" or "B" or "C"
	 * @param		startExcelRow	This is the third parameter to specify the excel row where data starts.<br>
	 * 					Exp. 1 or 2 or 3
	 */
	public void createFooter(Sheet sheet, String startExcelColumn,int startExcelRow);
}
