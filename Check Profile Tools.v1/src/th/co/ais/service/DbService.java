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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.commons.lang3.exception.ExceptionUtils;
import org.apache.log4j.Logger;

import th.co.ais.dto.ExcelSheetHead;
public class DbService {
	private static final Logger logger = Logger.getLogger(DbService.class);
	Connection conn=null;
	
	public DbService() {}
	
	public DbService(Connection conn) {
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

	public void insertToInvGenBatchAct(ExcelSheetHead excelSheetHead) {
		PreparedStatement pst = null;
		try {
			StringBuilder sql = new StringBuilder();
			sql.append("INSERT INTO CC_TBL_DAT_INV_GEN_BATCH_ACT ");
			sql.append(" (INVOICE_NUM, ACCOUNT_NUM, MOBILE, WAIVE, request_desc,item_desc,  DESCRIPTION,CUSTOMER_REF,SUB_MAIL,  main_cause,sub_cause,service_code) ");
			sql.append(" VALUES(?,?,?,?,?,?,?,?,?,?,?,?)");
			pst = conn.prepareStatement(sql.toString());
			
			pst.setString(1, excelSheetHead.getInvoiceNum());
			pst.setString(2, excelSheetHead.getAccountNum());
			pst.setString(3, excelSheetHead.getMobile());
			pst.setString(4, excelSheetHead.getWaive());
			pst.setString(5, excelSheetHead.getRequestDesc());
			pst.setString(6, excelSheetHead.getItemDesc());
			pst.setString(7, excelSheetHead.getDescription());
			pst.setString(8, excelSheetHead.getCustomerRef());
			pst.setString(9, excelSheetHead.getSubMail());
			pst.setString(10, excelSheetHead.getMainCause());
			pst.setString(11, excelSheetHead.getSubCause());
			pst.setString(12, excelSheetHead.getServiceCode());
			logger.info(ReflectionToStringBuilder.toString(excelSheetHead));
			pst.executeUpdate();
			
		} catch (Exception e) {
			logger.error("Insert insertToInvGenBatchAct Exception ",e);
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
