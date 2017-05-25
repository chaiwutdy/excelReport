package com.report.util;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;

public class FontUtils {

	public static Font blackBold(Sheet sheet){
		Font font = sheet.getWorkbook().createFont();
    font.setColor(HSSFColor.BLACK.index);
    font.setBold(true);
    return font;
	}
	
	public static Font whiteBold(Sheet sheet){
		Font font = sheet.getWorkbook().createFont();
    font.setColor(HSSFColor.WHITE.index);
    font.setBold(true);
    return font;
	}
}
