package com.report.dao.impl;

import java.util.Date;
import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.report.dao.BackOrderRptDAO;
import com.report.model.BackOrderSub1Model;
import com.report.model.BackOrderSub2Model;
import com.report.model.BackOrderSub3Model;

public class BackOrder3MRptDAOImpl extends JdbcDaoSupport implements BackOrderRptDAO {
	
	private static final Logger LOGGER = LogManager.getLogger(BackOrder3MRptDAOImpl.class.getName());

	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<BackOrderSub1Model> getSubReport1Data(Date reportDate) {

		List<BackOrderSub1Model> result = null;
		try {
			StringBuffer sql = new StringBuffer();
			LOGGER.info(sql.toString());
			result = getJdbcTemplate().query(sql.toString(), new BeanPropertyRowMapper(BackOrderSub1Model.class));
		} catch (Exception ex) { 
			ex.printStackTrace();
			LOGGER.info(ex.getMessage());
		}
		return result;
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<BackOrderSub2Model> getSubReport2Data(Date reportDate) {
		List<BackOrderSub2Model> result = null;
		try {
			StringBuffer sql = new StringBuffer();
			LOGGER.info(sql.toString());
			result = getJdbcTemplate().query(sql.toString(), new BeanPropertyRowMapper(BackOrderSub2Model.class));
		} catch (Exception ex) { 
			ex.printStackTrace();
			LOGGER.info(ex.getMessage());
		}
		return result;
	}
	
	
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<BackOrderSub3Model> getSubReport3Data(Date reportDate) {
		List<BackOrderSub3Model> result = null;
		try {
			StringBuffer sql = new StringBuffer();
			
			LOGGER.info(sql.toString());
			result = getJdbcTemplate().query(sql.toString(), new BeanPropertyRowMapper(BackOrderSub3Model.class));
		} catch (Exception ex) { 
			ex.printStackTrace();
			LOGGER.info(ex.getMessage());
		}
		return result;
	}

}
