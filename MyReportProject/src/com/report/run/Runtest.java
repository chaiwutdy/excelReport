package com.report.run;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	
public class Runtest {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		writeIntoExcel("AdvanceSituationReport.xlsx");
	}

	public static void writeIntoExcel(String file) throws FileNotFoundException, IOException{
        Workbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Sheet1");	
        Row row = sheet.createRow(0); 

        Cell name = row.createCell(0);
        name.setCellValue("Gokul");
        
        Cell birthdate = row.createCell(1);
        DataFormat format = book.createDataFormat();
        CellStyle dateStyle = book.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
        birthdate.setCellStyle(dateStyle);

        birthdate.setCellValue(new Date(115, 10, 10));
        
        sheet.autoSizeColumn(1);
        
        book.write(new FileOutputStream(file));
        book.close();
    }

	/*private Workbook getWorkbook(String excelFilePath)
	        throws IOException {
	    Workbook workbook = null;
	 
	    if (excelFilePath.endsWith("xlsx")) {
	        workbook = new XSSFWorkbook();
	    } else if (excelFilePath.endsWith("xls")) {
	        workbook = new HSSFWorkbook();
	    } else {
	        throw new IllegalArgumentException("The specified file is not Excel file");
	    }
	 
	    return workbook;
	}*/

	/*XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("template.xlsx"));           
    FileOutputStream fileOut = new FileOutputStream("new.xlsx");
    //Sheet mySheet = wb.getSheetAt(0);
    XSSFSheet sheet1 = wb.getSheet("Summary");
    XSSFRow row = sheet1.getRow(15);
    XSSFCell cell = row.getCell(3);
    cell.setCellValue("Bharthan");

    wb.write(fileOut);
    log.info("Written xls file");
    fileOut.close();*/
	
	
	/*Connection conn = null;	
    try {
        if (booLog) LogHelper.doInfo(this, "GetHistoryLocationNG (" + FrameNo + ")");
        conn = util.openConnection(conn);
        TNdcmaintransactionFacadeModified tNdcmaintransactionFacade = new TNdcmaintransactionFacadeModified();
        HashMap<String, String> hmCond = new HashMap<String, String>();
        hmCond.put(tNdcmaintransactionFacade.STATUS, IConstant.STATUS_OPEN);
        List<TNdcmaintransactionModel> tNdcmaintransactionModels = tNdcmaintransactionFacade.findByKeyAndHistoryLocationNG(conn, hmCond, FrameNo);
        if (tNdcmaintransactionModels != null && tNdcmaintransactionModels.size() > 0) {
            result = tNdcmaintransactionModels.get(0).getLocationid();
        }
    } catch (SQLException e) {
        util.rollbackConnection(conn);
        result = e.getMessage();
        LogHelper.doError(this, "GetHistoryLocationNG", e);
    } catch (Exception e) {
        util.rollbackConnection(conn);
        result = e.getMessage();
        LogHelper.doError(this, "GetHistoryLocationNG", e);
    } finally {
        util.closeConnection(conn);
    }*/
}
