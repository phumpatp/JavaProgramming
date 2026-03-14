package th.co.ais.service;

import java.io.BufferedWriter;
import java.math.BigDecimal;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Types;
import java.text.SimpleDateFormat;
//import java.time.LocalDateTime;  // Import the LocalDateTime class
//import java.time.format.DateTimeFormatter;  // Import the DateTimeFormatter class

import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
//import java.sql.Date;
import java.util.Date;

import javax.swing.JOptionPane;

import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.log4j.Logger;

import th.co.ais.dto.ExcelInputTemplate;
import th.co.ais.dto.ExcelSheetHeadAdjust;
public class DbServiceCheckProfileTools {
	private static final Logger logger = Logger.getLogger(DbServiceCheckProfileTools.class);
	Connection conn=null;
	
	public DbServiceCheckProfileTools() {}
	
	public DbServiceCheckProfileTools(Connection conn) {
		this.conn = conn;
	}
	
	public void deleteInvGenBatchAct(){
		PreparedStatement pst = null;
		try {
			StringBuilder sql = new StringBuilder();
			sql.append("DELETE FROM CC_TBL_DAT_INV_GEN_BATCH_ACT ");

			pst = conn.prepareStatement(sql.toString());
			

			pst.executeUpdate();
			
		} catch (Exception e) {
			logger.error("Delete deleteInvGenBatchAct Exception ",e);
			JOptionPane.showMessageDialog(null, "Delete deleteInvGenBatchAct Exception :: Error"+e);
		       
		}finally {
			closeStatement(pst,null);
		}
	}

	/*public void checkProfile5G(ExcelInputTemplate excelSheetHead) {
		PreparedStatement pst = null;
		try {
			String strSql;
			strSql="select distinct count(event_source) v_count from custeventsource " +
					" where customer_ref in (select customer_ref " +
					" from account where account_num=? " +
					" and event_type_id=1 and event_source=? ";

			PreparedStatement stmt=conn.prepareStatement(strSql);
			stmt.setString(1, accountNum);
			stmt.setString(2,mobileNumber);
			ResultSet rs=stmt.executeQuery();
		    int count=rs.getInt("v_count");
		    System.out.println(count);
		    conn.close();
		    if (count>0) {
		    	return true;
		    }else {
		    	JOptionPane.showConfirmDialog(null, "not have mobile active in period");
		    	return false;
		    }
		} catch (Exception e) {
			return false;
		}
	}
	*/
	private void closeStatement(Statement st, ResultSet rs){
		try{
			if(st != null){
				st.close();
			}
			if(rs != null){
				rs.close();
			}
		}catch (Exception e) {
			logger.error("Closes Statement Error",e);
		}
	}
	
	
	private static String removeNull(Object object){
		if(object == null || "null".equals(object)){
			return "";
		}else{
			return String.valueOf(object);
		}		
	}
	
	private void closeStatement(Statement st, ResultSet rs, BufferedWriter bw){
		try{
			if(st != null){
				st.close();
			}
			if(rs != null){
				rs.close();
			}
			if(bw != null){
				bw.close();
			}
		}catch (Exception e) {
			logger.error("Closes Statement Error"+ExceptionUtils.getStackTrace(e));
		}
	}
	
}
