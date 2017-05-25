package com.report.rpt;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.dao.BackOrderRptDAO;
import com.report.dao.CommonDAO;
import com.report.dao.StockRptDAO;
import com.report.generator.ReportGenerator;
import com.report.rpt.criteria.ReportCriteria;
import com.report.rpt.source.BackOrderMainSource;
import com.report.rpt.source.BackOrderSub1Source;
import com.report.rpt.source.BackOrderSub2Source;
import com.report.rpt.source.BackOrderSub3Source;
import com.report.rpt.source.ReportSource;
import com.report.rpt.source.StockMainSource;
import com.report.rpt.source.StockSub1Source;
import com.report.rpt.source.StockSub2Source;
import com.report.rpt.source.StockSub3Source;
import com.report.util.ExcelUtils;
import com.report.util.Utils;

public class NDODailySalesRpt implements ReportGenerator{
	
	private static final Logger LOGGER = LogManager.getLogger(NDODailySalesRpt.class.getName());
	
	@Override
	public String getFilePath(ReportCriteria reportCriteria) {
		String filePath = Utils.getProperties("app.path.ndoDailySalesRpt");
		return Utils.getFilePathWithDateFormat(filePath, reportCriteria.getReportDate());
	}

	@Override
	public Workbook getWorkbook(ReportCriteria reportCriteria) {
		Workbook book = new XSSFWorkbook();		
		//callProcedureForGetNewData();
		createBackOrderSheet(book, reportCriteria);
		createStockSheet(book, reportCriteria);
		LOGGER.info("Check Import Data Error: SELECT * FROM LOG_ERR_WUTTEST ORDER BY CREATE_DT DESC;");
		return book;
	}
	
	private void callProcedureForGetNewData(){
		CommonDAO commonDAO = (CommonDAO)Utils.APP_CONTEXT.getBean("commonDAO");
		commonDAO.callProcedure("CALL STOCK_DAILY_REPORT_WUTTEST2()");
		commonDAO.callProcedure("CALL BACK_ORDER_RPT_WUTTEST()");
	}
	
	private void createBackOrderSheet(Workbook book, ReportCriteria criteria){
		BackOrderRptDAO backOrderRptDAO = (BackOrderRptDAO)Utils.APP_CONTEXT.getBean("backOrderRptDAO");
		BackOrderRptDAO backOrder3MRptDAO = (BackOrderRptDAO)Utils.APP_CONTEXT.getBean("backOrder3MRptDAO");
		BackOrderRptDAO backOrderMore3MRptDAO = (BackOrderRptDAO)Utils.APP_CONTEXT.getBean("backOrderMore3MRptDAO");
		
		createNDODailySalesSheet(book, "Total BO", criteria, 
				new BackOrderMainSource(null), 
				new BackOrderSub1Source(backOrderRptDAO), 
				new BackOrderSub2Source(backOrderRptDAO), 
				new BackOrderSub3Source(backOrderRptDAO));

		createNDODailySalesSheet(book, "Total BO (3M)", criteria, 
				new BackOrderMainSource("Aging within 3 month"), 
				new BackOrderSub1Source(backOrder3MRptDAO), 
				new BackOrderSub2Source(backOrder3MRptDAO), 
				new BackOrderSub3Source(backOrder3MRptDAO));
	
		createNDODailySalesSheet(book, "Total BO (More 3M)", criteria, 
				new BackOrderMainSource("Aging morethan 3 month"), 
				new BackOrderSub1Source(backOrderMore3MRptDAO), 
				new BackOrderSub2Source(backOrderMore3MRptDAO), 
				new BackOrderSub3Source(backOrderMore3MRptDAO));
	}
	 
	private void createStockSheet(Workbook book, ReportCriteria criteria){
		StockRptDAO stockRptDAO = (StockRptDAO)Utils.APP_CONTEXT.getBean("stockRptDAO");
		StockRptDAO stockAdvanceRptDAO = (StockRptDAO)Utils.APP_CONTEXT.getBean("stockAdvanceRptDAO");
		StockRptDAO stockNormalRptDAO = (StockRptDAO)Utils.APP_CONTEXT.getBean("stockNormalRptDAO");
		
		createNDODailySalesSheet(book, "Total Stock", criteria, 
				new StockMainSource("Total Stock (Normal & Advance)"), 
				new StockSub1Source(stockRptDAO),
				new StockSub2Source(stockRptDAO), 
				new StockSub3Source(stockRptDAO));

		createNDODailySalesSheet(book, "Total Stock (Advance)", criteria, 
				new StockMainSource("Total Stock (Advance)"), 
				new StockSub1Source(stockAdvanceRptDAO), 
				new StockSub2Source(stockAdvanceRptDAO), 
				new StockSub3Source(stockAdvanceRptDAO));

		createNDODailySalesSheet(book, "Total Stock (Normal)", criteria, 
				new StockMainSource("Total Stock (Normal)"), 
				new StockSub1Source(stockNormalRptDAO), 
				new StockSub2Source(stockNormalRptDAO), 
				new StockSub3Source(stockNormalRptDAO));
	}
	
	private void createNDODailySalesSheet(Workbook book,String sheetName, ReportCriteria criteria,
			ReportSource main, ReportSource sub1, ReportSource sub2, ReportSource sub3){
		
		Sheet sheet = book.createSheet(sheetName);
		main.createHeader(sheet, "A", 3, criteria);
		
		//sub1
		int startSub1RowNum = ExcelUtils.getLastRow(sheet)+4;
		sub1.createHeader(sheet, "D", startSub1RowNum, criteria);	
		int startBodySub1RowNum = ExcelUtils.getLastRow(sheet)+1;
		sub1.createBody(sheet, "D", startBodySub1RowNum, criteria);
		sub1.createFooter(sheet,  "D", ExcelUtils.getLastRow(sheet)+1, criteria, startBodySub1RowNum);
		
		//sub2
		int startSub2RowNum = ExcelUtils.getLastRow(sheet)+3;
		sub2.createHeader(sheet, "D", startSub2RowNum, criteria);
		int startBodySub2RowNum = ExcelUtils.getLastRow(sheet)+1;
		sub2.createBody(sheet, "D", startBodySub2RowNum, criteria);
		sub2.createFooter(sheet,  "D", ExcelUtils.getLastRow(sheet)+1, criteria, startBodySub2RowNum);
		
		//sub3
		int startSub3RowNum = ExcelUtils.getLastRow(sheet)+3;
		sub3.createHeader(sheet, "D", startSub3RowNum, criteria);
		int startBodySub3RowNum = ExcelUtils.getLastRow(sheet)+1;
		sub3.createBody(sheet, "A", startBodySub3RowNum, criteria);
		sub3.createFooter(sheet, "D", ExcelUtils.getLastRow(sheet)+1, criteria, startBodySub3RowNum);
	}
	
}
