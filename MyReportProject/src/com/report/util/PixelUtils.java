package com.report.util;

public class PixelUtils {
	 public static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
	 public static final int UNIT_OFFSET_LENGTH = 7;
	 public static final int[] UNIT_OFFSET_MAP = new int[] { 0, 36, 73, 109, 146, 182, 219 }; 
    /*public static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
    public static final short EXCEL_ROW_HEIGHT_FACTOR = 20;
    public static final int UNIT_OFFSET_LENGTH = 7;
    public static final int[] UNIT_OFFSET_MAP = new int[] { 0, 36, 73, 109, 146, 182, 219 };
   
    public static short pixel2WidthUnits(int pxs) {
        short widthUnits = (short) (EXCEL_COLUMN_WIDTH_FACTOR * (pxs / UNIT_OFFSET_LENGTH));
        widthUnits += UNIT_OFFSET_MAP[(pxs % UNIT_OFFSET_LENGTH)]; 
        return widthUnits;
    }
    
    public static int widthUnits2Pixel(short widthUnits) {
        int pixels = (widthUnits / EXCEL_COLUMN_WIDTH_FACTOR) * UNIT_OFFSET_LENGTH;
        int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
        pixels += Math.floor((float) offsetWidthUnits / ((float) EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));  
        return pixels;
    }
    
    public static int heightUnits2Pixel(short heightUnits) {
        int pixels = (heightUnits / EXCEL_ROW_HEIGHT_FACTOR);
        int offsetWidthUnits = heightUnits % EXCEL_ROW_HEIGHT_FACTOR;
        pixels += Math.floor((float) offsetWidthUnits / ((float) EXCEL_ROW_HEIGHT_FACTOR / UNIT_OFFSET_LENGTH));  
        return pixels;
    }*/
  /**
   * excel width units(units of 1/256th of a character width) to pixel units
   * @param widthUnits
   * @return
   */
 /* public static int widthUnits2Pixel(short widthUnits) {
    int pixels = (widthUnits * EXCEL_COLUMN_WIDTH_FACTOR);
    System.out.println(pixels);
  
    int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
    pixels += Math.round((float) offsetWidthUnits / ((float) EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));
    return pixels;
  } */
	
	public static int widthUnits2Pixel(float widthUnits) {
		//return (widthUnits * 256) + 150;//200;
		return Math.round(widthUnits * 256) + 150;
	}
} 