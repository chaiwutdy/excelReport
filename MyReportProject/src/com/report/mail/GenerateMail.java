package com.report.mail;

import java.util.List;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.report.constant.ReportType;
import com.report.util.Utils;

public class GenerateMail {
	private static final Logger LOGGER = LogManager.getLogger(GenerateMail.class.getName());

	public void send(ReportType pReportType, String fileDataSourceName){

		MailGenerator mail = getMail(pReportType);
		try {
			if(mail != null){
				
				mail.setDataOfMail(fileDataSourceName);
				String subject = mail.getSubject();
				List<String> toList = mail.getMailToList();
				List<String> ccList = mail.getMailCCList();
				Multipart content = mail.getContent();
				
				MimeMessage message = createDraft(subject, toList, ccList, content);
				Transport.send(message);
				LOGGER.info("Message Send.....");
			}else{
				LOGGER.info("ReportType Not Found");
			}
		} catch (MessagingException e) {
			e.printStackTrace();
		}
		
	}
	
	private MimeMessage createDraft(String subject, List<String> toList ,List<String> ccList, Multipart content) throws AddressException, MessagingException{
		String host = Utils.getProperties("mail.host");
		String from = Utils.getProperties("mail.from");
    Properties properties = System.getProperties();

    properties.setProperty("mail.smtp.host", host);

    Session session = Session.getDefaultInstance(properties);

    MimeMessage message = new MimeMessage(session);

    message.setFrom(new InternetAddress(from));
    
    if(toList != null){
	    for(String to: toList){
	    	message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
	    }
    }
    
    if(ccList != null){
	    for(String cc: ccList){
	    	message.addRecipient(Message.RecipientType.CC, new InternetAddress(cc));
	    }
    }
    
    message.setSubject(subject);
    message.setContent(content);
    return message;
	}
	
	private MailGenerator getMail(ReportType pReportType){
		for(ReportType reportType:ReportType.values()){
			if(reportType.getId() == pReportType.getId()){
				return reportType.getMailGenerator();
			}
		}
		return null;
	}
	
}
