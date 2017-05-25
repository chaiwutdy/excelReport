package com.report.mail;

import java.util.ArrayList;
import java.util.List;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.internet.MimeBodyPart;

import com.report.constant.ReportType;
import com.report.dao.MailDAO;
import com.report.rpt.criteria.ReportCriteria;
import com.report.util.DateUtils;
import com.report.util.Utils;

public class NDODailySalesMail implements MailGenerator {
	
	private String subject;
	private List<String> toList;
	private List<String> ccList;
	private Multipart mp;
	
	@Override
	public List<String> getMailToList() {
		return toList;
	}

	@Override
	public List<String> getMailCCList() {
		return ccList;
	}

	@Override
	public String getSubject() {
		return subject;
	}

	@Override
	public Multipart getContent() {
		return mp;
	}

	@Override
	public void setDataOfMail(String fileDataSourceName) throws MessagingException {
		MimeBodyPart mbp1 = new MimeBodyPart();
    MailDAO mailDAO = (MailDAO)Utils.APP_CONTEXT.getBean("mailDAO");
		ReportCriteria reportCriteria = ReportType.NDO_DAILY_SALES_RPT.getReportCriteria();
    String reportDate = DateUtils.getDateByformat(reportCriteria.getReportDate(),"dd-MM-yyyy");
    
		subject = "NDO DAILY SALES(STK,BO) as of " + reportDate;
		
    FileDataSource fds = new FileDataSource(fileDataSourceName);
    if (fds.getFile().length() != 0) {
    	subject = "[Autosent] "+ subject;
    	toList = mailDAO.getToList();
    	ccList = mailDAO.getCCList();
    	String msgText1 = "*** Confidential *** Please do not  forward this e-mail.";
      mbp1.setText(msgText1);
      
      MimeBodyPart mbp2 = new MimeBodyPart();
      mbp2.setDataHandler(new DataHandler(fds));
      mbp2.setFileName(fds.getName());
      
      mp.addBodyPart(mbp2);
      mp.addBodyPart(mbp1);
      
    }else{
    	subject = "[Autosent ERROR] " + subject;
    	toList = new ArrayList<String>();
    	toList.add("xxx@xxx.xxx");
    	String etext = "Please check Log File in D:\\xxx";
		  mbp1.setText(etext);
		  mp.addBodyPart(mbp1);
    }
		
	}

}
