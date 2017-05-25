package com.report.util;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * This class is about utility for creating excel file by Apache Poi.<br>
 * Understanding about cell position of Apache Poi.<br>
 * There is difference cell position between Apache Poi And Excel.<br>
 * These are some example about it<br>
 * <table border = 1>
 * <tr>
 * 	<td>Position</td>		
 * 	<td>Apache Poi</td>		
 * 	<td>Excel</td>	
 * </tr>
 * 
 * <tr>
 *  <td>Row</td>		
 * 	<td>0</td>		
 * 	<td>1</td>	
 * </tr>
 * 
 * <tr>
 * 	<td>&nbsp;</td>		
 * 	<td>1</td>		
 * 	<td>2</td>	
 * </tr>
 * 
 * <tr>
 * 	<td>Column</td>		
 * 	<td>0</td>		
 * 	<td>A</td>	
 * </tr>
 * 
 * <tr>
 * 	<td>&nbsp;</td>		
 * 	<td>1</td>		
 * 	<td>B</td>		
 * </tr>
 * </table>
 * So when refer about cell position this class try to receive parameters and return value with excel position.
 */
public class ExcelUtils {

	/**
	 * This method to create a cell.
	 * @param		excelColumn This is the first parameter to specify excel column.<br>
	 * 					Exp. "A" or "B" or "C"
	 * @param		excelRow This is the second parameter to specify excel row.<br>
	 * 					Exp. 1 or 2 or 3
	 * @param		sheet	This is the third parameter to use create a cell.
	 * @return	cellObject.
	 */
	public static Cell getCell(String excelColumn, int excelRow, Sheet sheet){
		String cellReference = excelColumn+excelRow;
    return getCell(cellReference, sheet);
	}
	
	/**
	 * This method to create a cell.
	 * @param		cellReference This is the first parameter to specify cellReference in Excel.<br>
	 * 					Exp. "A1" or "B1" or "C1"
	 * @param		sheet	This is the second parameter to use create a cell.<br>
	 * @return	cellObject.
	 */
	public static Cell getCell(String cellReference, Sheet sheet){
		CellReference cellRef = new CellReference(cellReference);
		Row row = sheet.getRow(cellRef.getRow());
		if(sheet.getRow(cellRef.getRow()) == null){
			row = sheet.createRow(cellRef.getRow()); 
		}
    return row.createCell(cellRef.getCol());
	}
	
	/**
	 * This method to use only this class.
	 * @param		rowIndex This is the first parameter to specify the rowIndex of cell.
	 * @param		rowIndex This is the second parameter to specify the columnIndex of cell.
	 * @param		sheet	This is the third parameter to use create a cell.
	 * @return	cellObject.
	 */
	private static Cell getCell(int rowIndex,int columnIndex, Sheet sheet){
		Row row = sheet.getRow(rowIndex);
		if(sheet.getRow(rowIndex) == null){
			row = sheet.createRow(rowIndex); 
		}
    return row.createCell(columnIndex);
	}
	
	/**
	 * This method to get a previousColumnCell.
	 * @param		cell This is the parameter to reference the currentCell to get the previousColumn.
	 * @return	cellObject of previousColumn.
	 */
	public static Cell getPreviousColumn(Cell cell){
		return getCell(cell.getRowIndex(), cell.getColumnIndex()-1, cell.getSheet());
	}
	
	/**
	 * This method to get a nextColumnCell.
	 * @param		cell This is the parameter to reference the currentCell to get the nextColumn.
	 * @return	cellObject of nextColumn.
	 */
	public static Cell getNextColumn(Cell currentCell){
		return getCell(currentCell.getRowIndex(), currentCell.getColumnIndex()+1, currentCell.getSheet());
	}
	
	/**
	 * This method to get a previousRowCell.
	 * @param		cell This is the parameter to reference the currentCell to get the previousRow.
	 * @return	cellObject of previousRow.
	 */
	public static Cell getPreviousRow(Cell currentCell){
		return getCell(currentCell.getRowIndex()-1, currentCell.getColumnIndex(), currentCell.getSheet());
	}
	
	/**
	 * This method to get a nextRowCell.
	 * @param		cell This is the parameter to reference the currentCell to get the nextRowCell.
	 * @return	cellObject of nextRow.
	 */
	public static Cell getNextRow(Cell currentCell){
		return getCell(currentCell.getRowIndex()+1, currentCell.getColumnIndex(), currentCell.getSheet());
	}
	
	/**
	 * This method for merging Columns.
	 * @param		sheet	This is the first parameter to use mergeColumn.
	 * @param		firstCell	This is the second parameter to specify start of position.
	 * @param		lastCell This is the third parameter to specify end of position.
	 */
	public static void mergeColumn(Sheet sheet, Cell firstCell, Cell lastCell){
		sheet.addMergedRegion(new CellRangeAddress(firstCell.getRowIndex(),
													firstCell.getRowIndex(),
													firstCell.getColumnIndex(),
													lastCell.getColumnIndex()));
	}
	
	/**
	 * This method for merging Rows.
	 * @param		sheet	This is the first parameter to use merge rows.
	 * @param		firstCell	This is the second parameter to specify the start of merge position.
	 * @param		lastCell This is the third parameter to specify the end of merge position.
	 */
	public static void mergeRow(Sheet sheet, Cell firstCell, Cell lastCell){
		sheet.addMergedRegion(new CellRangeAddress(firstCell.getRowIndex(),
													lastCell.getRowIndex(),
													firstCell.getColumnIndex(),
													firstCell.getColumnIndex()));
	}
	
	/**
	 * This method for merging Cells.
	 * @param		sheet	This is the first parameter to use merge cells.
	 * @param		firstCell	This is the second parameter to specify the start of merge position.
	 * @param		lastCell This is the third parameter to specify the end of merge position.
	 */
	public static void mergeCell(Sheet sheet,  Cell firstCell, Cell lastCell){
		sheet.addMergedRegion(new CellRangeAddress(firstCell.getRowIndex(),
													lastCell.getRowIndex(),
													firstCell.getColumnIndex(),
													lastCell.getColumnIndex()));
	}
	
	/**
	 * This method always return lastRow of lastData in excel.
	 * @param		sheet	This is the first parameter to use get a last row.
	 * @return	lastRow of lastData in excel.
	 */
	public static int getLastRow(Sheet sheet){
		return sheet.getLastRowNum()+1;
	}
	
	/**
	 * This method to create the dummyColumn.
	 * @param		startCell	This is the first parameter that reference to start to create the dummyColumn.
	 * @param		cellStlye	This is the second parameter to set cellStlye each the dummyColumn.
	 * @param		columnLength This is the third parameter to specify how long the dummyColumn.
	 */
	public static void createDummyColumn(Cell startCell, CellStyle cellStlye, int columnLength){
		Cell dummyCell = startCell;
		for(int i=1;i<=columnLength;i++){
			dummyCell = ExcelUtils.getNextColumn(dummyCell);
			dummyCell.setCellStyle(cellStlye);
		}	
	}
	
	/**
	 * This method to create the dummyRow.
	 * @param		startCell	This is the first parameter that reference to start to create the dummyRow.
	 * @param		cellStlye	This is the second parameter to set cellStlye each the dummyRow.
	 * @param		columnLength This is the third parameter to specify how long the dummyRow.
	 */
	public static void createDummyRow(Cell startCell, CellStyle cellStlye, int rowLength){
		Cell dummyCell = startCell;
		for(int i=1;i<=rowLength;i++){
			dummyCell = ExcelUtils.getNextRow(dummyCell);
			dummyCell.setCellStyle(cellStlye);
		}	
	}
}
