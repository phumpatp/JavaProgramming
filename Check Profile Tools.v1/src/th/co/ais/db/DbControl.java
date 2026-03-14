package th.co.ais.db;

import java.io.File;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Properties;

import org.apache.log4j.Logger;

import javax.swing.*;
import java.awt.*;


public class DbControl {
	private static final Logger logger = Logger.getLogger(DbControl.class);
	
	 public Connection getConnectionOra(String username,String password,String database) {
	        Connection conn;
	        try {
	            DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
	    		//System.out.println(args[1]);
	    		//System.out.println(args[2]);
	            //JOptionPane.showMessageDialog(null, "Connect Database :: Error"+"xx"+username+"xx"+database);
	            if ("RBMSIT".equalsIgnoreCase(database)){
		            conn = DriverManager.getConnection(
		            		"jdbc:oracle:thin:@(DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 10.104.192.45)(PORT = 1532)))(CONNECT_DATA = (SERVER = DEDICATED) (SERVICE_NAME = rbmsit)))",
		            		username,
		            		password);
		            logger.info("Connect Database :: Success ");
	            }
	            else if ("RBMPROD".equalsIgnoreCase(database)){
	            	//JOptionPane.showMessageDialog(null, "Connect Database :: RBMPROD");
		            conn = DriverManager.getConnection(
		            		"jdbc:oracle:thin:@10.232.66.76:1521:RBMPROD",
		            		username,
		            		password);

		            logger.info("Connect Database :: Success ");
		            //JOptionPane.showMessageDialog(null, "Connect Database :: Done");
	            } else if ("rbmprod_cirro".equalsIgnoreCase(database)){
	            	//JOptionPane.showMessageDialog(null, "Connect Database :: CIRRO");
		            conn = DriverManager.getConnection(
		            		"jdbc:oracle:thin:@(DESCRIPTION=    (ADDRESS=      (PROTOCOL=TCP)      (HOST=10.235.121.121)      (PORT=1521)    )    (CONNECT_DATA=      (SERVICE_NAME=rbmprod)    )  )",
		            		username,
		            		password);

		            logger.info("Connect Database :: Success ");
		            //JOptionPane.showMessageDialog(null, "Connect Database :: Done");
	            }else if ("rbmprod2_cirro".equalsIgnoreCase(database)){
	            	//JOptionPane.showMessageDialog(null, "Connect Database :: CIRRO");
		            conn = DriverManager.getConnection(
		            		"jdbc:oracle:thin:@(DESCRIPTION =    (ADDRESS_LIST =    (ADDRESS = (PROTOCOL = TCP)(HOST = 10.232.66.76)(PORT = 1521))    )    (CONNECT_DATA =      (SERVER = DEDICATED)      (SERVICE_NAME = RBMPROD)    )  )",
		            		username,
		            		password);

		            logger.info("Connect Database :: Success ");
		            //JOptionPane.showMessageDialog(null, "Connect Database :: Done");
	            }
	            else {
	            	conn=null;
	            	logger.error("Connect Database :: Don't have database");
	            }
 
	            return conn;
	        } catch (Exception e) {
	            e.printStackTrace();
	            logger.error("Connect Database :: Error ", e);
	            JOptionPane.showMessageDialog(null, "Connect Database :: Error"+database+e);
	            conn=null;
	            //return null;
	            
	        }
	        return conn;
	    }
	
	public Properties loadGlobalConfig(String configFile) throws Exception{
		File file = new File(configFile);
		Properties prop = new Properties();
		logger.info("Load properties config "+configFile);
		try {
			prop.load(new FileInputStream(file));
			logger.info("Load Global properties config success ");
		} catch (Exception e) {
			logger.error("Load Global properties config fail ",e);
			throw e;
		}	
		return prop;
	}	

	
	public void closeStatement(Statement statement){
		if(statement!=null){
			try {
				 statement.close();
			} catch (SQLException e) {
				logger.error("closeStatement Fail ",e);
				e.printStackTrace();
			}
		} 
		
	}
	
	public void closeConnection(Connection connection){
		if(connection!=null){
			try {
				connection.close();
			} catch (SQLException e) {
				logger.error("closeConnection Fail ",e);
				e.printStackTrace();
			}
		}
	}
	
	public void closeResultSet(ResultSet resultSet){
		if(resultSet!=null){
			try {
				resultSet.close();
			} catch (SQLException e) {
				logger.error("closeResultSet Fail ",e);
				e.printStackTrace();
			}
		} 
	}
	
	private void closeStatement(Connection con, Statement st, ResultSet rs){
		try{
			if(con!=null) {
				con.close();
			}
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
}
