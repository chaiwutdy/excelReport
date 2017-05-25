package com.report.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

public class FormulasUtils {

	public static String sum(Cell cell,int startRow,int endRow){
		String colRef = CellReference.convertNumToColString(cell.getColumnIndex());
		return "SUM("+colRef+startRow+":"+colRef+endRow+")";
	}
	
	public static String subtract(Cell firstCell,Cell... nextCells){
		String formulas = null;
		CellReference nextCellRef = null;
		CellReference firstCellRef = new CellReference(firstCell);
		formulas = firstCellRef.formatAsString();
		for(Cell nextcell:nextCells){
			nextCellRef =  new CellReference(nextcell);
			formulas+="-"+nextCellRef.formatAsString();
		}
		return formulas;
	}
	
	public static String plus(Cell firstCell,Cell nextCells){
		CellReference firstCellRef = new CellReference(firstCell);
		CellReference nextCellRef = new CellReference(nextCells);
		return firstCellRef.formatAsString()+"+"+nextCellRef.formatAsString();
	}
	
}
