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

import th.co.ais.dto.ExcelSheetHeadAdjust;
public class DbServiceAdjust {
	private static final Logger logger = Logger.getLogger(DbServiceAdjust.class);
	Connection conn=null;
	
	public DbServiceAdjust() {}
	
	public DbServiceAdjust(Connection conn) {
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

	public void insertToadjustmentLoad(ExcelSheetHeadAdjust excelSheetHead) {
		PreparedStatement pst = null;
		try {
			StringBuilder sql = new StringBuilder();
			//SimpleDateFormat formatter = new SimpleDateFormat("dd/MMM/yyyy", Locale.ENGLISH);

			String dateInString;
			//Date date;
			
			sql.append("INSERT INTO cc_tbl_Dat_inv_adjustment_load ");
			sql.append(" (Account_num, Mobile_num, Adjustment_date, Adjustment_type, Revenue_Code_name, Adjustment_Amt, Activity_flag,Activity_TH, Activity_EN, SMS_flag, SMS_TH, SMS_EN, Bill_TH, Bill_EN,Adjustment_flag, Update_date, User_Created,Case_ID, Process_name)");
			sql.append(" VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,SYSDATE,'bl_admin',?,'1_LOAD')");
			pst = conn.prepareStatement(sql.toString());
			
			//sql.append(" (Account_num, Mobile_num, Adjustment_date, Adjustment_type, Revenue_Code_name, ;
			pst.setString(1, excelSheetHead.getaccountNum());
			pst.setString(2, excelSheetHead.getmobileNumber());
			//date = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(excelSheetHead.getadjustmentDate());
		   SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy", Locale.ENGLISH);
	       java.util.Date today = format.parse(excelSheetHead.getadjustmentDate());
	       pst.setDate(3, new java.sql.Date(today.getTime()));
	            
			//System.out.println(today);
			//Date parsedDate = sdf.parse(excelSheetHead.getadjustmentDate());
			pst.setString(4, excelSheetHead.getadjustmentType());
			pst.setString(5, excelSheetHead.getrevenueCodeName());
			//Adjustment_Amt, Activity_flag,Activity_TH, Activity_EN, SMS_flag, SMS_TH, SMS_EN, Bill_TH, 
			pst.setInt(6, new Integer(excelSheetHead.getadjustmentAmt())*1000);
			pst.setString(7, excelSheetHead.getactivityFlag());
			pst.setString(8, excelSheetHead.getactivityTH());
			pst.setString(9, excelSheetHead.getactivityEn());
			pst.setString(10, excelSheetHead.getsmsFlag());
			pst.setString(11, excelSheetHead.getsmsTH());
			pst.setString(12, excelSheetHead.getsmsEn());
			pst.setString(13, excelSheetHead.getbillTH());
			//Bill_EN,Adjustment_flag, Update_date, User_Created,Case_ID, Process_name")
			pst.setString(14, excelSheetHead.getbillEn());
			//date=formatter.parse(excelSheetHead.getupdateDate().toString());
			pst.setString(15, excelSheetHead.getadjustmentFlag());
			//logger.error( excelSheetHead.getadjustmentFlag());
			//pst.setString(16, timeStamp);
			//pst.setString(16, excelSheetHead.getuserCreated());
			pst.setString(16, excelSheetHead.getcaseID());
			//pst.setString(18, excelSheetHead.getprocessName());

			//logger.error(sql.toString()+pst.toString());	
			//logger.info(ReflectionToStringBuilder.toString(excelSheetHead));
			pst.executeUpdate();
			
		} catch (Exception e) {
			logger.error("Insert insertToAdjustment_Load Exception ",e);
			JOptionPane.showMessageDialog(null, "Insert insertToInvGenBatchAct Exception :: Error"+e);
		    
		}finally {
			closeStatement(pst,null);
		}
	}
	
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
