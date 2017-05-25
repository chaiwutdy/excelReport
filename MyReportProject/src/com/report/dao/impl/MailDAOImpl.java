package com.report.dao.impl;

import java.util.List;

import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.report.dao.MailDAO;

public class MailDAOImpl extends JdbcDaoSupport implements MailDAO {
	
	@Override
	public List<String> getToList() {
		
		return null;
	}

	@Override
	public List<String> getCCList() {
		
		return null;
	}
	

}
