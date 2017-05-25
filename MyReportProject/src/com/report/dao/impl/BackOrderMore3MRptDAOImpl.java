package com.report.dao.impl;

import java.util.Date;
import java.util.List;

import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.report.dao.BackOrderRptDAO;
import com.report.model.BackOrderSub1Model;
import com.report.model.BackOrderSub2Model;
import com.report.model.BackOrderSub3Model;

public class BackOrderMore3MRptDAOImpl extends JdbcDaoSupport implements BackOrderRptDAO {

	@Override
	public List<BackOrderSub1Model> getSubReport1Data(Date reportDate) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<BackOrderSub2Model> getSubReport2Data(Date reportDate) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<BackOrderSub3Model> getSubReport3Data(Date reportDate) {
		// TODO Auto-generated method stub
		return null;
	}

}
