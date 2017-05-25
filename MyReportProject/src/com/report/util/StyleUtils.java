package com.report.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class StyleUtils {

	public static CellStyle allBorder(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderYellow(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderYellow2(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		XSSFColor color = new XSSFColor(new java.awt.Color(252,252,144));
		((XSSFCellStyle) style).setFillForegroundColor(color);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderYellowAlignCenter(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		XSSFColor color = new XSSFColor(new java.awt.Color(252,252,144));
		((XSSFCellStyle) style).setFillForegroundColor(color);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderRedAlignCenter(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		//style.setFillForegroundColor(IndexedColors.ROSE.getIndex());
		XSSFColor color = new XSSFColor(new java.awt.Color(240,199,152));
		((XSSFCellStyle) style).setFillForegroundColor(color);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderBlueAlignCenter(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
//		style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
		XSSFColor color = new XSSFColor(new java.awt.Color(44,119,218)); 
		((XSSFCellStyle) style).setFillForegroundColor(color);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderAlignCenter(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setWrapText(true);
    return style;
	}
	
	public static CellStyle allBorderAlignRight(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setWrapText(true);
    return style;
	}
	
	
	public static CellStyle allBorderGrayAlignCenter(CellStyle style){
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
//		style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
		XSSFColor color = new XSSFColor(new java.awt.Color(202,192,192)); 
		((XSSFCellStyle) style).setFillForegroundColor(color);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setWrapText(true);
    return style;
	}
	
}
