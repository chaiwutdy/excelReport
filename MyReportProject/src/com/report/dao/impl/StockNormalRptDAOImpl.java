package com.report.dao.impl;

import java.util.Date;
import java.util.List;

import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.report.dao.StockRptDAO;
import com.report.model.StockSub1Model;
import com.report.model.StockSub2Model;
import com.report.model.StockSub3Model;

public class StockNormalRptDAOImpl extends JdbcDaoSupport implements StockRptDAO {

	@Override
	public List<StockSub1Model> getSubReport1Data(Date reportDate) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<StockSub2Model> getSubReport2Data(Date reportDate) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<StockSub3Model> getSubReport3Data(Date reportDate) {
		// TODO Auto-generated method stub
		return null;
	}
	

}
