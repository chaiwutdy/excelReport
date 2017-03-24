package com.report.util;

import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.Locale;

public class DateUtils {

	private static final Locale localeTH = new Locale("th", "TH");
//    private static final Locale localeEN = Locale.US;
    
	public static String getCurrentDate(String format) {
    Date currentDate = new Date();
    SimpleDateFormat sdf = new SimpleDateFormat(format, localeTH);
    return sdf.format(currentDate);
  }
	
	public static String[] getShortMonths(){
		String [] months =  new DateFormatSymbols(Locale.US).getShortMonths();
		return Arrays.copyOf(months, months.length-1);
	}
}
