package com.report.dao.impl;

import java.util.List;

import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.report.dao.InvAdvRptSitDAO;
import com.report.model.ChannelModel;
import com.report.model.MastCarSeriesModel;
import com.report.rowmapper.MastCarSeriesRowMapper;

public class InvAdvRptSitDAOImpl extends JdbcDaoSupport implements InvAdvRptSitDAO{

	public List<MastCarSeriesModel> getSubReport1Data() {
		List<MastCarSeriesModel> result = null;
		try {
			StringBuffer sql = new StringBuffer();
			sql.append("SELECT SERIES_ENAME, DISPLAY_SEQ ");
			sql.append("FROM MAST_CAR_SERIES_WOOT2 ");
			sql.append("WHERE DISPLAY_SEQ IS NOT NULL " );
			sql.append("ORDER BY DISPLAY_SEQ " );
			result = getJdbcTemplate().query(sql.toString(), new MastCarSeriesRowMapper());
		} catch (Exception ex) { 
			ex.printStackTrace();
		}
		return result;
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public List<ChannelModel> getSubReport2Data() {
		List<ChannelModel> result = null;
		try {
			StringBuffer sql = new StringBuffer();
			sql.append("SELECT DISTINCT ZONE_ID, AREA_ID ");
			sql.append("FROM CHANNEL ");
			sql.append("WHERE STATUS = 'A' " );
			sql.append("AND DEALER_FLAG IN ('3','1') " );
			sql.append("ORDER BY ZONE_ID, AREA_ID " );
			result = getJdbcTemplate().query(sql.toString(), new BeanPropertyRowMapper(ChannelModel.class));
		} catch (Exception ex) { 
			ex.printStackTrace();
		}
		return result;
	}
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public List<ChannelModel> getSubReport3Data() {
		List<ChannelModel> result = null;
		try {
			StringBuffer sql = new StringBuffer();
			sql.append("SELECT DISTINCT ZONE_ID, AREA_ID, CHANNEL_CODE, ");
			sql.append("CH_SHORT_NAME, DISPLAY_SEQ ");
			sql.append("FROM CHANNEL ");
			sql.append("WHERE STATUS = 'A' " );
			sql.append("AND DEALER_FLAG IN ('3','1') " );
			sql.append("ORDER BY DISPLAY_SEQ " );
			result =getJdbcTemplate().query(sql.toString(), new BeanPropertyRowMapper(ChannelModel.class));
		} catch (Exception ex) { 
			ex.printStackTrace();
		}
		return result;
	}
	
}
