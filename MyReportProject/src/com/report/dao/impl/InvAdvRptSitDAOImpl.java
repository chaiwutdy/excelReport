package com.report.dao.impl;

import java.util.List;

import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.report.dao.InvAdvRptSitDAO;
import com.report.model.ChannelModel;
import com.report.model.MastCarSeriesModel;

public class InvAdvRptSitDAOImpl extends JdbcDaoSupport implements InvAdvRptSitDAO{

	@Override
	public List<MastCarSeriesModel> getSubReport1Data() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<ChannelModel> getSubReport2Data() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<ChannelModel> getSubReport3Data() {
		// TODO Auto-generated method stub
		return null;
	}

	
}
