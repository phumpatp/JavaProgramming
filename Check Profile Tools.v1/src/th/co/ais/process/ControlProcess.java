package th.co.ais.process;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.security.spec.ECFieldF2m;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.OldExcelFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import com.sun.tools.javac.util.Convert;

import th.co.ais.db.DbControl;
import th.co.ais.dto.ExcelInputTemplate;
import th.co.ais.dto.ExcelSheetHead;
import th.co.ais.dto.ExcelSheetHeadAdjust;
import th.co.ais.service.DbService;
import th.co.ais.service.DbServiceAdjust;
import th.co.ais.service.DbServiceCheckProfileTools;

import javax.sound.sampled.AudioFormat.Encoding;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;


public class ControlProcess {
private static final Logger logger = Logger.getLogger(ControlProcess.class);
	
	public static void main(String[] args) {
		
		try {
			logger.info("check profile tools start");
			//int result;
			String filename = null;
			String pathInput=null;
			int j;
			FileDialog dialog = new FileDialog((Frame)null, "Select File to Open");
			dialog.setFile("*.xls;*.xlsx;*.txt;*.GMF;*cut*;*.sql");
		    dialog.setMode(FileDialog.LOAD);
		    dialog.setVisible(true);
		    //dialog.setFile(".xls");
		    String filenameinput = dialog.getFile();
		    pathInput=dialog.getDirectory();
		    dialog.dispose();
		    System.out.println(filenameinput + " chosen." + pathInput);
		   
			String extension = "";
			//filename=xxx.xls
			//filename=xx
			//extension=xls
		      final JTextField username = new JTextField(5);
		      //JPasswordField password = new JPasswordField(5);
		      JTextField password = new JTextField(5);
		      JTextField database = new JTextField(5);
		      JRadioButton rb3= new JRadioButton() ;
		      JRadioButton rb5= new JRadioButton() ;
		      JRadioButton rb6= new JRadioButton() ;  
		      
		      JButton b;   
			if ((filenameinput!=null)) {
			 j = filenameinput.lastIndexOf('.');
			if (j >= 0) {
			    extension = filenameinput.substring(j+1);
			    filename=filenameinput.substring(j);
			    //System.out.println(extension);
			    //System.out.println(filenameinput.length());
			}

		    //if ( (filenameinput.length()>=1) && (("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension))))
			if ( (filenameinput.length()>=1) )
		    {
		    	//System.out.println("xxx");

			      int result=0;
			  if (("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension))){
	
		      
		      JPanel myPanel = new JPanel();
		      //myPanel.setLayout(new GridLayout(1, 0));
		      myPanel.setLayout(new BoxLayout(myPanel, BoxLayout.Y_AXIS));
		      //myPanel.setLayout(new BorderLayout(myPanel,,BorderLayout.PAGE_END));
		      myPanel.add(new JLabel("username:"));
		      myPanel.add(username,BorderLayout.LINE_START);
		      
		      myPanel.add(Box.createHorizontalStrut(10)); // a spacer
		      myPanel.add(new JLabel("password:"));
		      myPanel.add(Box.createHorizontalStrut(10));
		      myPanel.add(password,BorderLayout.LINE_END);
		      //password.setEchoChar('*');
		      myPanel.add(Box.createHorizontalStrut(10)); // a spacer
		      myPanel.add(new JLabel("database:"));
		      myPanel.add(database,BorderLayout.LINE_START);
		      myPanel.add(Box.createHorizontalStrut(10));
		      myPanel.add(new JLabel("\n"));
		      myPanel.add(new JLabel("----------Select Report--------------------"));
		      myPanel.add(new JLabel("\n"));
		      myPanel.add(Box.createHorizontalStrut(10));
		        
		      myPanel.setAlignmentX(Component.RIGHT_ALIGNMENT);
		      
		       rb3=new JRadioButton("Service 5G");    
		       rb3.setBounds(100,150,80,30);
		          
		       rb5=new JRadioButton("Check Promotion");    
		        
		       rb6=new JRadioButton("Check Promotion FBB");    
		      
		      ButtonGroup bg =new ButtonGroup();
		      bg.add(rb3);
		      rb3.setBounds(300, 20, 300, 25);
		      myPanel.add(rb3,BorderLayout.PAGE_END);
		      //myPanel.add(Box.createHorizontalStrut(200));
		      
		      bg.add(rb5);
		      myPanel.add(rb5,BorderLayout.PAGE_END);
		      
		      bg.add(rb6);
		      myPanel.add(rb6,BorderLayout.PAGE_END);
		      //myPanel.add(bg);
		       
		      myPanel.setAlignmentX(Component.LEFT_ALIGNMENT);
		      myPanel.add(new JLabel("\n"));
		      
		      
		      username.addActionListener(new ActionListener() {
		          @Override
		          public void actionPerformed(ActionEvent e) {
		              final String text = username.getText();
		             System.out.println("test");
		          }
		      });
		      
		      JOptionPane pane = new JOptionPane(myPanel, JOptionPane.QUESTION_MESSAGE, JOptionPane.OK_CANCEL_OPTION) {
		            @Override
		            public void selectInitialValue() {
		                username.requestFocusInWindow();
		            }
		        };
		      
		      result = JOptionPane.showConfirmDialog(null, myPanel, 
		               "Please Enter username/password/database", JOptionPane.OK_CANCEL_OPTION);
			  }

		    //if ( username.getText(.gen)!=0 && password.length()!=0 && database.length()!=0) { 
			if ((filenameinput.contains("Input_Template"))) {
				//JOptionPane.showMessageDialog(null, "test file 5G");
				//try {

				if ("*" !=extension) {
					//try {
					  if ((result == JOptionPane.OK_OPTION)) {
						
				    	if ( username.getText().length()!=0 && password.getText().length()!=0 && database.getText().length()!=0) {
				    	   if (rb3.isSelected()) {
				    		   System.out.println("5g");
				    		   processCheckService5G(filenameinput, filename, extension, pathInput,username.getText(),password.getText(),database.getText());
				    	   }else if (rb5.isSelected()) {
				    		   System.out.println("check promotion");
				    		   processCheckPromotion(filenameinput, filename, extension, pathInput,username.getText(),password.getText(),database.getText());
				    	   }else if (rb6.isSelected()) {
				    		   System.out.println("check promotion fbb");
				    		   processCheckPromotionFBB(filenameinput, filename, extension, pathInput,username.getText(),password.getText(),database.getText());
				    	   } else {
				    		   JOptionPane.showMessageDialog(null, "please select one report");
				    	   }
				    	
				    	}
				    	else {
				    		JOptionPane.showMessageDialog(null, "please check username/password or database");
				    	}
				    }
				}
			}else		{
				JOptionPane.showMessageDialog(null, "file not support");
			}

			logger.info("check profile tools end");
		   
		    }
			}else
			{
				JOptionPane.showMessageDialog(null, "please check file name");
	    	}
		} catch (Exception e) {
			//if (result != JOptionPane.CANCEL_OPTION) {
			//if (filenameinput)
			logger.error("check profile Tools error", e);
		    //JOptionPane.showMessageDialog(null, "load file error" +e);
			//}
		}
		
	}
	
	
	public static void splitBilldata(String filenameinput,String filename,String extension,String pathInput,String accountnum) throws IOException
	{
		String line;
        BufferedReader masterfile;
        String tag[]=null;
        int lineCount=0;
        String outfilename=null;
        
        int numberaccount =0;
        
        String Accountnum=null;
       
        String strEncoding=null;
      //outfilename=Accountnum+"_out.gmf";
       
       //System.setProperty("file.edcoding","UTF-8");
       //masterfile = new BufferedReader(new FileReader(pathInput+filenameinput));
	  /*    JTextField inAccountnum = new JTextField(5);
	      int result=0;
	      
        JPanel myPanel = new JPanel();
	      myPanel.setLayout(new BoxLayout(myPanel, BoxLayout.Y_AXIS));
	      myPanel.add(new JLabel("Accountnum:"));
	      myPanel.add(inAccountnum);
	      myPanel.add(Box.createHorizontalStrut(15)); // a spacer
	     myPanel.setAlignmentX(Component.LEFT_ALIGNMENT);

	     result = JOptionPane.showConfirmDialog(null, myPanel, 
	               "Please Enter account for cut", JOptionPane.OK_CANCEL_OPTION);
	     if (result!=JOptionPane.CANCEL_OPTION) {*/
        Accountnum= accountnum;
       //Accountnum="31989071711002";
       masterfile=  new BufferedReader(
       new InputStreamReader(new FileInputStream(pathInput+filenameinput)));

       //strEncoding=System.getProperty("master.encoding");
       
       line = masterfile.readLine();
       int intDocStart=0;
       int intAccountNo=0;
       int intDocEnd=0;
       
       String indexfile ;
       FileWriter fw;
       BufferedWriter bw;
   	   
       File tempfile;
       boolean goloop=true;
       try {
        while((line != null)&& goloop)
        {
        	   // System.out.println(line);
        	    tag=line.split(" ");
        	    lineCount++;
        	  
    		   if (line.toString().length()>0) {
      	        //if (tag[0].contains("DOCSTART"))  {
        	    if (line.contains("DOCSTART")) {
	        	
	        	   //write file temp
      	        	intDocStart=lineCount;
	        	   }
      	        if (tag[0].equalsIgnoreCase("ACCOUNTNO") && tag[1].equalsIgnoreCase(Accountnum)) {
      	        	//System.out.println("acccountno" +tag[1]);
      	        	 intAccountNo=lineCount;
      	        	 //System.out.println("acccountno" +intAccountNo);
      	        	 line=masterfile.readLine();
      	        	 tag=line.split(" ");
      	        	 goloop =true;
      	        	 while(goloop) {
      	        		 lineCount++;
      	        		 //System.out.println("acccountno" +lineCount);
      	        		 line=masterfile.readLine();
      	        		 if (line==null) {
      	        			 goloop=false;
      	        		 }else if( line.toString().length()>0) {
      	        			//System.out.println("acccountnoxx" +line);
      	        			 if ( line.contains("DOCSTART")){
	      	        			intDocEnd=lineCount;
	      	        			goloop=false;
	      	        			break;
      	        			 }
      	        			 
      	        		 } else if (line.contains("DOCEND")) {
	      	        			intDocEnd=lineCount;
	      	        			//System.out.println("acccountnoxx" +line);
	      	        			goloop=false;
	      	        			break;
      	        		 }
      	        	 }
      	        	//System.out.println("acccountno" +lineCount);
      	        	 
      	        }
		        //if (tag[0].contains("DOCEND"))  {
      	      //System.out.println("acccountnoyy" +lineCount);
      	        if (line.contains("DOCEND")) {

		    	   //write file temp
			        intDocEnd=lineCount;
		    	   } 
    		   }//end if lenght
               line = masterfile.readLine();
               if (line.toString().length()==0) {
            	   break;
               }
             
        }
       } catch (Exception e){
    	   intDocEnd=lineCount;
    	   //System.out.println("acccountnoyy" +lineCount);
       }
        masterfile.close();
        System.out.println("docstart_ " +intDocStart);
        System.out.println("accountno " + intAccountNo);
        System.out.println("docend " +intDocEnd);
        //System.out.println(lineCount);
        masterfile.close();
        //write to new file
        if (intAccountNo>0) {
         // masterfile = new BufferedReader(new FileReader(pathInput+filenameinput));
         masterfile = new BufferedReader(new InputStreamReader(new FileInputStream(pathInput+filenameinput),"tis620"));

         line = masterfile.readLine();
 
    	 //Accountnum="32000050255040";
    	 lineCount=0;
    	 //indexfile=Integer.toString(numberaccount);
		    outfilename=Accountnum + "_out.GMF";
			   
		    tempfile = new File(pathInput+outfilename);
	   
	        fw = new FileWriter(tempfile.getAbsoluteFile());
	        //bw = new BufferedWriter(fw);
	        bw=new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tempfile),"tis620"));
	        if (!tempfile.exists()) {
	        	tempfile.createNewFile();
   	        }
   	         else {
   	        	tempfile.delete();
   	        	tempfile.createNewFile();
   	         }
	     lineCount=0;
         while(line != null)
         {
         	   //System.out.println(line);
         	    //System.out.println("OK");
         	    tag=line.split(" ");
         	    lineCount++;
         	   //if (tag[0].contains("DOCSTART")) {
     		   //numberaccount++;
         	  
      		   if (lineCount==intDocStart) {
      			  //System.out.println(line);
      			   bw.write(line);
      			   bw.write("\n");
      			   //bw.newLine();
      	
     			   while(((lineCount<intDocEnd) )){
     				   line=masterfile.readLine();
     				   lineCount++;
     				   //System.out.println(line);
     				   bw.write(line);
     				   bw.write("\n");
          			   //bw.newLine();
     			   }
     			   bw.close();
     			   break;
     			  
     		   }	
                line = masterfile.readLine();
         }//while
         
         masterfile.close();
        } else
        {
        	 System.out.println("not found account");
        }
       
	// }//end if 
	}
	
	public static void removeAccountfromBilldata(String filenameinput,String filename,String extension,String pathInput,String accountnum) throws IOException
	{
		String line;
        BufferedReader masterfile;
        String tag[]=null;
        int lineCount=0;
        String outfilename=null;
        
        int numberaccount =0;
        
        String Accountnum=null;
       
        String strEncoding=null;
      //outfilename=Accountnum+"_out.gmf";
       
       //System.setProperty("file.edcoding","UTF-8");
       //masterfile = new BufferedReader(new FileReader(pathInput+filenameinput));
	   /*   JTextField inAccountnum = new JTextField(5);
	      int result=0;
	      
        JPanel myPanel = new JPanel();
	      myPanel.setLayout(new BoxLayout(myPanel, BoxLayout.Y_AXIS));
	      myPanel.add(new JLabel("Accountnum:"));
	      myPanel.add(inAccountnum);
	      myPanel.add(Box.createHorizontalStrut(15)); // a spacer
	     myPanel.setAlignmentX(Component.LEFT_ALIGNMENT);

	     result = JOptionPane.showConfirmDialog(null, myPanel, 
	               "Please Enter account for cut", JOptionPane.OK_CANCEL_OPTION);
	     if (result!=JOptionPane.CANCEL_OPTION) {*/
        Accountnum= accountnum;
       //Accountnum="31989071711002";
       masterfile=  new BufferedReader(
       new InputStreamReader(new FileInputStream(pathInput+filenameinput)));

       //strEncoding=System.getProperty("master.encoding");
       
       line = masterfile.readLine();
       int intDocStart=0;
       int intAccountNo=0;
       int intDocEnd=0;
       
       String indexfile ;
       FileWriter fw;
       BufferedWriter bw;
   	   
       File tempfile;
       
       boolean goloop=true;
       try {
        while((line != null)&& goloop)
        {
        	   // System.out.println(line);
        	    tag=line.split(" ");
        	    lineCount++;
        	  
    		   if (line.toString().length()>0) {
      	        //if (tag[0].contains("DOCSTART"))  {
        	    if (line.contains("DOCSTART")) {
	        	
	        	   //write file temp
      	        	intDocStart=lineCount;
	        	   }
      	        if (tag[0].equalsIgnoreCase("ACCOUNTNO") && tag[1].equalsIgnoreCase(Accountnum)) {
      	        	//System.out.println("acccountno" +tag[1]);
      	        	 intAccountNo=lineCount;
      	        	 //System.out.println("acccountno" +intAccountNo);
      	        	 line=masterfile.readLine();
      	        	 tag=line.split(" ");
      	        	 goloop =true;
      	        	 while(goloop) {
      	        		 lineCount++;
      	        		 //System.out.println("acccountno" +lineCount);
      	        		 line=masterfile.readLine();
      	        		 if (line==null) {
      	        			 goloop=false;
      	        		 }else if( line.toString().length()>0) {
      	        			//System.out.println("acccountnoxx" +line);
      	        			 if ( line.contains("DOCSTART")){
	      	        			intDocEnd=lineCount;
	      	        			goloop=false;
	      	        			break;
      	        			 }
      	        			 
      	        		 } else if (line.contains("DOCEND")) {
	      	        			intDocEnd=lineCount;
	      	        			//System.out.println("acccountnoxx" +line);
	      	        			goloop=false;
	      	        			break;
      	        		 }
      	        	 }
      	        	//System.out.println("acccountno" +lineCount);
      	        	 
      	        }
		        //if (tag[0].contains("DOCEND"))  {
      	      //System.out.println("acccountnoyy" +lineCount);
      	        if (line.contains("DOCEND")) {

		    	   //write file temp
			        intDocEnd=lineCount;
		    	   } 
    		   }//end if lenght
               line = masterfile.readLine();
               if (line.toString().length()==0) {
            	   break;
               }
             
        }
       } catch (Exception e){
    	   intDocEnd=lineCount;
    	   //System.out.println("acccountnoyy" +lineCount);
       }
        masterfile.close();
        System.out.println("docstart_ " +intDocStart);
        System.out.println("accountno " + intAccountNo);
        System.out.println("docend " +intDocEnd);
        //System.out.println(lineCount);
        masterfile.close();
        //write to new file
        if (intAccountNo>0) {
         // masterfile = new BufferedReader(new FileReader(pathInput+filenameinput));
         masterfile = new BufferedReader(new InputStreamReader(new FileInputStream(pathInput+filenameinput),"tis620"));

         line = masterfile.readLine();
 
    	 //Accountnum="32000050255040";
    	 lineCount=0;
    	 //indexfile=Integer.toString(numberaccount);
		    outfilename=filenameinput + "_new."+ extension;
			   
		    tempfile = new File(pathInput+outfilename);
	   
	        fw = new FileWriter(tempfile.getAbsoluteFile());
	        //bw = new BufferedWriter(fw);
	        bw=new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tempfile),"tis620"));
	        if (!tempfile.exists()) {
	        	tempfile.createNewFile();
   	        }
   	         else {
   	        	tempfile.delete();
   	        	tempfile.createNewFile();
   	         }
	     lineCount=0;
         while(line!=null)
         {
         	   //System.out.println(line);
         	    //System.out.println("OK");
        	    //masterfile.readLine();
         	    tag=line.split(" ");
         	    lineCount++;
         	   //if (tag[0].contains("DOCSTART")) {
     		   //numberaccount++;
         	  
      		   if ((lineCount <intDocStart)||(lineCount >intDocEnd)) {
      			  //System.out.println(line);
      			  bw.write(line);
      			  bw.write("\n");
     		   }	
                line = masterfile.readLine();
         }//while
         bw.close();
         masterfile.close();
        } else
        {
        	 System.out.println("not found account");
        }
	 //}//end if 
	}
	
	public static void processGenBatchAct(String filenameinput,String filename,String extension,String pathInput,String username,String password,String database) {
		//String pathInput = "C:\\xlsFloder/";
		//System.out.println(pathInput);
		org.apache.poi.ss.usermodel.Sheet sheet;
		//Sheet sheet;

		//JOptionPane.showConfirmDialog(null,filenameinput);

	    	  
			DbControl db = new DbControl();
			Connection conn = null;


			try {
				//if ((filename.length()>=0) ) {
		
					conn = db.getConnectionOra(username,password,database);
					DbService dbService = new DbService(conn);
					if (conn !=null) {
					//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format"+conn);
						
					dbService.deleteInvGenBatchAct();
					logger.info("delete before insert");
				       
					//	if(("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension)))  {
						//logger.info("read sheet :"+extension);
						logger.info("read file name :"+filenameinput);
						//JOptionPane.showMessageDialog(null, fileNameInput);
						//ArrayList<ExcelSheetHead> excelDataList = new ArrayList<ExcelSheetHead>();
						try {
							
							//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format");
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							org.apache.poi.ss.usermodel.Workbook wb=WorkbookFactory.create(ExcelFileToRead);
							
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
							for (int i = 2; i <= lastRow; i++) {
								ExcelSheetHead listRow = new ExcelSheetHead();
								Row rowCust = sheet.getRow(i);
								if(rowCust != null) {
									logger.info("row :"+(i-1));
									listRow.setInvoiceNum(getCellDataCust(rowCust.getCell(0)));
									listRow.setAccountNum(getCellDataCust(rowCust.getCell(1)));
									listRow.setMobile(getCellDataCust(rowCust.getCell(2)));
									listRow.setWaive(getCellDataCust(rowCust.getCell(3)));
									listRow.setRequestDesc(getCellDataCust(rowCust.getCell(4)));
									listRow.setItemDesc(getCellDataCust(rowCust.getCell(5)));
									listRow.setDescription(getCellDataCust(rowCust.getCell(6)));
									listRow.setCustomerRef(getCellDataCust(rowCust.getCell(7)));
									listRow.setSubMail(getCellDataCust(rowCust.getCell(8)));
									listRow.setSubCause(getCellDataCust(rowCust.getCell(9)));
									listRow.setMainCause(getCellDataCust(rowCust.getCell(10)));
									listRow.setServiceCode(getCellDataCust(rowCust.getCell(11)));
									//logger.info(ReflectionToStringBuilder.toString(listRow));
									dbService.insertToInvGenBatchAct(listRow);
								}
								
							}
							logger.info("new excel format ");
							
						} catch( OLE2NotOfficeXmlFileException e){
							JOptionPane.showMessageDialog(null, "old excel HSSFWorkbook format" +e);
							//JOptionPane.showMessageDialog(null, "  new excel format");
							
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
							//org.apache.poi.ss.usermodel.Workbook wb=WorkbookFactory.create(ExcelFileToRead);
							

							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
							logger.info("old excel format ");
							
						}
						JOptionPane.showMessageDialog(null, "load file done");
					//}
					
					//}  for loop
					} //else
					//{
					//	JOptionPane.showMessageDialog(null, "connection error");
					//}
				//}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				//if (result!=JOptionPane.CANCEL_OPTION) {
			
					logger.error(e);
					//JOptionPane.showMessageDialog(null, "load file error");
					JOptionPane.showMessageDialog(null, "load file :: Error"+e);
					//JOptionPane.showMessageDialog(null, "old excel format" +e);
			           
				//}
			}
	}//end sub
	
	public static String CheckRevenueCode(String adjustmentType,String adjustmmentAmt) {
		
		//'Dim strAdjustmentType As String
	    //' cn 45   bg -45 0000
	    //' dn -45  bg +45 1111
		double adjamt=0;
		String getRevenueCode=null;
		adjamt= Double.parseDouble(adjustmmentAmt);  
		if (adjamt <0) {
		    //If Cells(lngRow, enuCol.adjustment_amt) < 0 Then
			if ("IR".equalsIgnoreCase(adjustmentType)){
				getRevenueCode="ADJ-IR-1111";
			} else if ("IDD".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-IDD-1111";
			}else if ("DATA".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-DT-1111";
			}else if ("AIRTIME".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-AT-1111";
			}else if ("GOOD".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-GS-1111";
			}
		}else {
			if ("IR".equalsIgnoreCase(adjustmentType)){
				getRevenueCode="ADJ-IR-0000";
	
			} else if ("IDD".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-IDD-0000";
			}else if ("DATA".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-DT-0000";
			}else if ("AIRTIME".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-AT-0000";
			}else if ("GOOD".equalsIgnoreCase(adjustmentType)) {
				getRevenueCode="ADJ-GS-0000";
			}
		}
		return getRevenueCode;
	}
	
	public static boolean checkAccount(Connection conn,String accountNum) {
		try {
			String strSql;
			int v_count=0;
			strSql="select count(*) v_count from account where account_num=?";
			PreparedStatement stmt=conn.prepareStatement(strSql);
			 System.out.println(strSql);
			stmt.setString(1, accountNum);
			ResultSet rs=stmt.executeQuery();
			while (rs.next()) {
			    v_count=rs.getInt("v_count");
			    System.out.println(v_count);
			    
			}
			conn.close();
		    if (v_count>0) {
		    	return true;
		    }else {
		    	JOptionPane.showConfirmDialog(null, "not have account active in period");
		    	return false;
		    }
		} catch (Exception e) {
			System.out.println(e);
			return false;
		}
	}
	
	public static boolean checkMobileActive(Connection conn,String accountNum,String mobileNumber) {
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
	
	public static void processAdjustmentLoad(String filenameinput,String filename,String extension,String pathInput,String username,String password,String database) {
		//String pathInput = "C:\\xlsFloder/";
		//System.out.println(pathInput);
		org.apache.poi.ss.usermodel.Sheet sheet;
		//Sheet sheet;
		boolean v_check=true;
	    	  
			DbControl db = new DbControl();
			Connection conn = null;

			try {
				//if ((filename.length()>=0) ) {
		
					conn = db.getConnectionOra(username,password,database);
					DbServiceAdjust dbServiceadj = new DbServiceAdjust(conn);
					if (conn !=null) {
					//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format"+conn);
						
					//dbService.deleteInvGenBatchAct();
					//logger.info("delete before insert");
				       
					//	if(("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension)))  {
						//logger.info("read sheet :"+extension);
						logger.info("read file name :"+filenameinput);
						//JOptionPane.showMessageDialog(null, fileNameInput);
						//ArrayList<ExcelSheetHead> excelDataList = new ArrayList<ExcelSheetHead>();
						try {
							
							//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format");
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							org.apache.poi.ss.usermodel.Workbook wb=WorkbookFactory.create(ExcelFileToRead);
							
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
							//JOptionPane.showMessageDialog(null, sheet.getLastRowNum());
							for (int i = 2; i <= lastRow; i++) {
								ExcelSheetHeadAdjust listRow = new ExcelSheetHeadAdjust();
								Row rowCust = sheet.getRow(i);
								//JOptionPane.showMessageDialog(null, rowCust.getCell(0));
								if((rowCust != null) && (rowCust.getCell(0).toString().length() >0)) {
									logger.info("row :"+(i-1));
//sql.append(" (Account_num, Mobile_num, Adjustment_date, Adjustment_type, Revenue_Code_name, ;
									//v_check=(checkAccount(conn, getCellDataCust(rowCust.getCell(0))));
									v_check=true;
									if (v_check==true){
									//JOptionPane.showMessageDialog(null, rowCust.getCell(0));
									listRow.setaccountNum(getCellDataCust(rowCust.getCell(0)));
									//v_check=(checkMobileActive(conn, getCellDataCust(rowCust.getCell(0)), getCellDataCust(rowCust.getCell(1))));
									v_check=true;
									if (v_check==true) {
										listRow.setmobileNumber(getCellDataCust(rowCust.getCell(1)));
										listRow.setadjustmentDate(rowCust.getCell(2).toString());
										listRow.setadjustmentType(getCellDataCust(rowCust.getCell(3)));
										//JOptionPane.showConfirmDialog(null, getCellDataCust(rowCust.getCell(3)));
										listRow.setadjustmentAmt(getCellDataCust(rowCust.getCell(4)));
										//JOptionPane.showConfirmDialog(null, getCellDataCust(rowCust.getCell(4)));
										String revenueCodeName;
										revenueCodeName=CheckRevenueCode( getCellDataCust(rowCust.getCell(3)).toString(), getCellDataCust(rowCust.getCell(4)).toString());
										//JOptionPane.showConfirmDialog(null, revenueCodeName);
										listRow.setrevenueCodeName(revenueCodeName);
	//Adjustment_Amt, Activity_flag,Activity_TH, Activity_EN, SMS_flag, SMS_TH, SMS_EN, Bill_TH, 
										
										listRow.setactivityFlag(getCellDataCust(rowCust.getCell(5)));
										listRow.setactivityTH(getCellDataCust(rowCust.getCell(6)));
										listRow.setactivityEn(getCellDataCust(rowCust.getCell(7)));
										listRow.setsmsFlag(getCellDataCust(rowCust.getCell(8)));
										listRow.setsmsTH(getCellDataCust(rowCust.getCell(9)));
										listRow.setsmsEn(getCellDataCust(rowCust.getCell(10)));
										listRow.setbillTH(getCellDataCust(rowCust.getCell(11)));
	//Bill_EN,Adjustment_flag, Update_date, User_Created,Case_ID, Process_name")
										listRow.setbillEn(getCellDataCust(rowCust.getCell(12)));
										listRow.setadjustmentFlag(null);
										//listRow.setupdateDate(null);
										//listRow.setuserCreated("bl_admin");
										listRow.setcaseID(getCellDataCust(rowCust.getCell(13)));
										//Adjustment_flag	 Update_date	 User_Created	Case_ID	 Process_name
										//listRow.setprocessName("1_LOAD");
										//listRow.setprocessName(getCellDataCust(rowCust.getCell(19)));
										//logger.info(ReflectionToStringBuilder.toString(listRow));
										//logger.info("row :"+getCellDataCust(rowCust.getCell(16)));
										dbServiceadj.insertToadjustmentLoad(listRow);
										}//check mobile
									}//check account
								}
								else {
									break;
								}
							}
							logger.info("new excel format ");
							
						} catch(OldExcelFormatException e){
							JOptionPane.showMessageDialog(null, "read excel error " +e);
					
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
						
							logger.info("old excel format ");
							
						}
						JOptionPane.showMessageDialog(null, "load file done");
					//}
					
					//}  for loop
					}// else

				//}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				//if (result!=JOptionPane.CANCEL_OPTION) {
			
					logger.error(e);
					//JOptionPane.showMessageDialog(null, "load file error");
					JOptionPane.showMessageDialog(null, "load file :: Error"+e);
					//JOptionPane.showMessageDialog(null, "old excel format" +e);
			           
				//}
			}
	}//end sub
	
	public static void processCheckService5G(String filenameinput,String filename,String extension,String pathInput,String username,String password,String database) {
		//String pathInput = "C:\\xlsFloder/";
		//System.out.println(pathInput);
		org.apache.poi.ss.usermodel.Sheet sheet;
		//org.apache.poi.ss.usermodel.Sheet sheetOut;
		//Sheet sheet;
		String strSql;
        String accountnum = null;
        String customerref = null;
        String tariffname = null;
        String productseqs = null;
        String  startdate = null;
        String enddate = null;
        String prodstatus = null;
        String mobilenum=null;
        String outputFile=null;
        String datenow =null;
		boolean v_check=true;
	    	  
			DbControl db = new DbControl();
			Connection conn = null;

			try {
				//if ((filename.length()>=0) ) {
		
					conn = db.getConnectionOra(username,password,database);
					DbServiceCheckProfileTools dbServicedbTools = new DbServiceCheckProfileTools(conn);
					if (conn !=null) {
					//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format"+conn);
						
					//dbService.deleteInvGenBatchAct();
					//logger.info("delete before insert");
				       
					//	if(("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension)))  {
						//logger.info("read sheet :"+extension);
						logger.info("read file name :"+filenameinput);
						//JOptionPane.showMessageDialog(null, fileNameInput);
						//ArrayList<ExcelSheetHead> excelDataList = new ArrayList<ExcelSheetHead>();
						try {
							
							//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format");
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							
							//read file
							org.apache.poi.ss.usermodel.Workbook wb=WorkbookFactory.create(ExcelFileToRead);
							
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
							//JOptionPane.showMessageDialog(null, sheet.getLastRowNum());
							
							//output file
							
							
						    /*SimpleDateFormat formatter = new SimpleDateFormat("YYYYMMDD");  
						    Date date = new Date();  
						    System.out.println(formatter.format(date));  
						    */
							
						    long millis=System.currentTimeMillis();  
						      
						    // creating a new object of the class Date  
						    java.sql.Date date = new java.sql.Date(millis);       
						    ///System.out.println(date);   
						    
						    outputFile="Check_Service_5G_Report_" + date + ".xlsx";
						    
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							//System.out.println("TEST");
							File fileName = new File(pathInput +outputFile);
						
					        FileOutputStream fos = new FileOutputStream(pathInput +outputFile);
					        XSSFWorkbook  workbook = new XSSFWorkbook();            

					        XSSFSheet sheetout = workbook.createSheet("service 5G");  
					
				         /*    worksheet_out.Cells[1, "A"] = "No.";
			                    worksheet_out.Cells[1, "B"] = "Mobile";
			                    worksheet_out.Cells[1, "C"] = "Status Mobile";
			                    worksheet_out.Cells[1, "D"] = "Service";
			                    worksheet_out.Cells[1, "E"] = "Product_Seq Service";
			                    worksheet_out.Cells[1, "F"] = "Start date";
			                    worksheet_out.Cells[1, "G"] = "End date";*/
			                    
			                int introw=0;
			                Row row = sheetout.createRow((short)introw);   
							Cell cell = row.createCell(0);
							cell.setCellValue("No.");
							row.createCell(1).setCellValue("mobile");
							row.createCell(2).setCellValue("Status Mobile");
							row.createCell(3).setCellValue("Service");
							row.createCell(4).setCellValue("Product_Seq Service");
							row.createCell(5).setCellValue("star_date");
							row.createCell(6).setCellValue("end_date");
							for (int i = 1; i <= lastRow; i++) {
								ExcelInputTemplate listRow = new ExcelInputTemplate();
								Row rowCust = sheet.getRow(i);
								//JOptionPane.showMessageDialog(null, rowCust.getCell(1));
								if((rowCust != null) && (rowCust.getCell(1).toString().length() >0)) {
									logger.info("row :"+(i-1));

									//JOptionPane.showMessageDialog(null, rowCust.getCell(0));
									//listRow.setaccountNum(getCellDataCust(rowCust.getCell(0)));
									listRow.setmoBile(getCellDataCust(rowCust.getCell(1)));
									//listRow.setproductSeq((rowCust.getCell(2)));
									//System.out.println(getCellDataCust(rowCust.getCell(1)));
									//dbServicedbTools.checkProfile5G(listRow);
									strSql= "   select distinct (select ce1.event_source from custeventsource ce1 where ce1.customer_ref=c.customer_ref   " +
											   "  and ce1.event_type_id=1 and ce1.event_source= ? " +
											   "  and rownum<=1) event_source,e.account_num,c.customer_ref, t.tariff_name,c.product_seq,  " +
											   "   to_char(d.START_DAT,'DD/MM/YYYY') start_dat, to_char(d.END_DAT,'DD/MM/YYYY') end_dat,c.parent_product_seq,  " +
											   "    (select decode(cpds.product_status,'OK','ACTIVE','SU','SUSPEND','TX','TERMINATE') product_status from custproductstatus cpds  " +
											   "   where cpds.customer_ref=c.customer_ref and cpds.product_seq=d.product_seq and cpds.effective_dtm =(select  min(cpds1.effective_dtm) from custproductstatus cpds1 where cpds1.customer_ref=cpds.customer_ref and cpds1.product_seq=cpds.product_seq and cpds1.effective_dtm between add_months(acc.next_bill_dtm,-1) and acc.next_bill_dtm-1 )) product_status   " +
											   "   from custhasproduct c,CUSTPRODUCTTARIFFDETAILS d,custproductdetails e, tariff t,  " +
											   "   CUSTEVENTSOURCE s,accountattributes aa ,account acc " +
											   "  where c.PRODUCT_SEQ=d.PRODUCT_SEQ  and acc.account_num=aa.account_num" +
											   "  and c.CUSTOMER_REF=d.CUSTOMER_REF and c.PRODUCT_SEQ=e.PRODUCT_SEQ  " +
											   "  and c.CUSTOMER_REF=e.CUSTOMER_REF  and d.TARIFF_ID = t.TARIFF_ID  " +
											   "  and (s.EVENT_TYPE_ID=1 or s.EVENT_TYPE_ID is null)  " +
											   "  and c.PRODUCT_SEQ=s.PRODUCT_SEQ(+) and c.CUSTOMER_REF=s.CUSTOMER_REF(+)  " +
											   "    and aa.ACCOUNT_NUM=e.ACCOUNT_NUM  " +
											   "   and aa.account_num= (select distinct account_num from custeventsource ce,custproductdetails cpd  " +
											   "  where ce.customer_ref = cpd.customer_ref  " +
											   "  and ce.event_source = ? " +
											   "   and ce.product_seq = cpd.product_seq  " +
											   "  and ce.event_type_id = 1  and rownum<=1)  " +
											   "   and c.parent_product_seq  in (select ce1.product_seq from custeventsource ce1  " +
											   "   where ce1.customer_ref=c.customer_ref   " + 
											   "  and ce1.event_type_id=1 and ce1.event_source=? " +
											   "   and rownum<=1)  " +
											   "   and t.tariff_name ='5G NSA Service' ";
									
									System.out.println(strSql);
									PreparedStatement stmt=conn.prepareStatement(strSql);
									mobilenum=getCellDataCust(rowCust.getCell(1));
									stmt.setString(1, getCellDataCust(rowCust.getCell(1)));
									stmt.setString(2, getCellDataCust(rowCust.getCell(1)));
									stmt.setString(3, getCellDataCust(rowCust.getCell(1)));
									ResultSet rs=stmt.executeQuery();
										
									//System.out.println("loop");
									while (rs.next()) {
									    //v_count=rs.getInt("v_count");
									       System.out.println(rs.getString("event_source"));
						                    accountnum = rs.getString("account_num");
				                            customerref = rs.getString("customer_ref");
				                            tariffname = rs.getString("tariff_name");
				                            productseqs = rs.getString("product_seq");
				                            startdate = rs.getString("start_dat");
				                            enddate = rs.getString("end_dat");
				                            prodstatus = rs.getString("product_status");
				                            
				                            //08xxxxxxxx
				                            //Active, Suspend etc.
				                            //5GNSA, 5SA
				                            //4271
				                            //09/06/2022 00:00
				                            //09/06/2024 00:00
				                            
									}
									introw++;
									//System.out.println(introw);
									rs.close();
									// Create a row and put some cells in it. Rows are 0 based.
									row = sheetout.createRow((short)introw);
									// Create a cell and put a value in it.
									row.createCell(0).setCellValue(introw);
									row.createCell(1).setCellValue(mobilenum);
									//row.createCell(2).setCellValue(accountnum);
									row.createCell(2).setCellValue(prodstatus);
									row.createCell(3).setCellValue(tariffname);
									row.createCell(4).setCellValue(productseqs);
									row.createCell(5).setCellValue(startdate);
									row.createCell(6).setCellValue(enddate);
									}//check account
							        

								    //System.out.println("out");
							}
							
							logger.info("new excel format ");
							workbook.write(fos);
					        fos.flush();
					        fos.close();
					        conn.close();
							
						} catch(OldExcelFormatException e){
							JOptionPane.showMessageDialog(null, "read excel error " +e);
					
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
						
							logger.info("old excel format ");
							
						}
						JOptionPane.showMessageDialog(null, "generate output file " + outputFile + " done");
					//}
					
					//}  for loop
					}// else

				//}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				//if (result!=JOptionPane.CANCEL_OPTION) {
			
					logger.error(e);
					//JOptionPane.showMessageDialog(null, "load file error");
					JOptionPane.showMessageDialog(null, "generate output file " + filename + "errr" +e);
					//JOptionPane.showMessageDialog(null, "old excel format" +e);
			           
				//}
			}
	}//end sub
	
	public static void processCheckPromotion(String filenameinput,String filename,String extension,String pathInput,String username,String password,String database) {
		//String pathInput = "C:\\xlsFloder/";
		//System.out.println(pathInput);
		org.apache.poi.ss.usermodel.Sheet sheet;
		//org.apache.poi.ss.usermodel.Sheet sheetOut;
		//Sheet sheet;
		String strSql;
        String accountnum = null;
        String customerref = null;
        String tariffname = null;
        String productseqs = null;
        String  startdate = null;
        String enddate = null;
        String prodstatus = null;
        String mobilenum=null;
        String outputFile=null;
        String datenow =null;
		boolean v_check=true;
	    	  
			DbControl db = new DbControl();
			Connection conn = null;

			try {
				//if ((filename.length()>=0) ) {
		
					conn = db.getConnectionOra(username,password,database);
					DbServiceCheckProfileTools dbServicedbTools = new DbServiceCheckProfileTools(conn);
					if (conn !=null) {
					//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format"+conn);
						
					//dbService.deleteInvGenBatchAct();
					//logger.info("delete before insert");
				       
					//	if(("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension)))  {
						//logger.info("read sheet :"+extension);
						logger.info("read file name :"+filenameinput);
						//JOptionPane.showMessageDialog(null, fileNameInput);
						//ArrayList<ExcelSheetHead> excelDataList = new ArrayList<ExcelSheetHead>();
						try {
							
							//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format");
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							
							//read file
							org.apache.poi.ss.usermodel.Workbook wb=WorkbookFactory.create(ExcelFileToRead);
							
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
							//JOptionPane.showMessageDialog(null, sheet.getLastRowNum());
							
							//output file
							
							
						    /*SimpleDateFormat formatter = new SimpleDateFormat("YYYYMMDD");  
						    Date date = new Date();  
						    System.out.println(formatter.format(date));  
						    */
							
						    long millis=System.currentTimeMillis();  
						      
						    // creating a new object of the class Date  
						    java.sql.Date date = new java.sql.Date(millis);       
						    ///System.out.println(date);   
						    
						    outputFile="Check_Promoton_Report_" + date + ".xlsx";
						    
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							//System.out.println("TEST");
							File fileName = new File(pathInput +outputFile);
						
					        FileOutputStream fos = new FileOutputStream(pathInput +outputFile);
					        XSSFWorkbook  workbook = new XSSFWorkbook();            

					        XSSFSheet sheetout = workbook.createSheet("Check_Promotion");  
					
				         /*     worksheet_out.Cells[1, "A"] = "No.";
			                    worksheet_out.Cells[1, "B"] = "Mobile";
			                    worksheet_out.Cells[1, "C"] = "BA No.";
			                    worksheet_out.Cells[1, "D"] = "Status Mobile";
			                    worksheet_out.Cells[1, "E"] = "Product Seq. Promotion";
			                    worksheet_out.Cells[1, "F"] = "Promotion_name";
			                    worksheet_out.Cells[1, "G"] = "Start date";
			                    worksheet_out.Cells[1, "H"] = "End date";;*/
			                    
			                int introw=0;
			                Row row = sheetout.createRow((short)introw);   
							Cell cell = row.createCell(0);
							cell.setCellValue("No.");
							row.createCell(1).setCellValue("mobile");
							row.createCell(2).setCellValue("BA No.");
							row.createCell(3).setCellValue("Status Mobile");
							row.createCell(4).setCellValue("Product Seq. Promotion");
							row.createCell(5).setCellValue("Promotion_name");
							row.createCell(6).setCellValue("star_date");
							row.createCell(7).setCellValue("end_date");
							for (int i = 1; i <= lastRow; i++) {
								ExcelInputTemplate listRow = new ExcelInputTemplate();
								Row rowCust = sheet.getRow(i);
								//JOptionPane.showMessageDialog(null, rowCust.getCell(1));
								if((rowCust != null) && (rowCust.getCell(1).toString().length() >0)) {
									logger.info("row :"+(i-1));

									//JOptionPane.showMessageDialog(null, rowCust.getCell(0));
									//listRow.setaccountNum(getCellDataCust(rowCust.getCell(0)));
									listRow.setmoBile(getCellDataCust(rowCust.getCell(1)));
									//listRow.setproductSeq((rowCust.getCell(2)));
									//System.out.println(getCellDataCust(rowCust.getCell(1)));
									//dbServicedbTools.checkProfile5G(listRow);
									strSql=" select distinct (select ce1.event_source from custeventsource ce1 " +
			                                 " where ce1.customer_ref=ce.customer_ref  " +
			                                 " and ce1.event_type_id=1 " +
			                                 " and rownum<=1) event_source, e.account_num,c.customer_ref, d.product_seq,t.tariff_name,to_char(d.START_DAT,'DD/MM/YYYY') start_dat, to_char(d.END_DAT,'DD/MM/YYYY') end_dat,  " +
			                                 " (select decode(cpds.product_status,'OK','ACTIVE','SU','SUSPEND','TX','TERMINATE') from custproductstatus cpds " +
			                                 " where cpds.customer_ref=c.customer_ref and cpds.product_seq=d.product_seq and cpds.effective_dtm =(select  min(cpds1.effective_dtm) from custproductstatus cpds1 where cpds1.customer_ref=cpds.customer_ref and cpds1.product_seq=cpds.product_seq and cpds1.effective_dtm between add_months(acc.next_bill_dtm,-1) and acc.next_bill_dtm-1 )) product_status  " +
			                                 " from custproductdetails e,account acc,CUSTPRODUCTTARIFFDETAILS d,tariff  t,custeventsource ce, " +
			                                 " custhasproduct c " +
			                                 " where e.customer_ref=acc.customer_ref " +
			                                 " and e.account_num=acc.account_num " +
			                                 " and e.customer_ref=d.customer_Ref " +
			                                 " and  c.PRODUCT_SEQ=d.PRODUCT_SEQ " +
			                                 " and c.CUSTOMER_REF=d.CUSTOMER_REF " +
			                                 " and c.PRODUCT_SEQ=e.PRODUCT_SEQ " +
			                                 " and c.CUSTOMER_REF=e.CUSTOMER_REF " +
			                                 " and d.tariff_id=t.tariff_id " +
			                                 " and acc.account_num= ? " +
			                                 " and ce.event_source= ? " +
			                                 " and d.product_seq= ? " +
			                                 " and ce.customer_ref=e.customer_ref " ;
			                                
									
									System.out.println(strSql);
									PreparedStatement stmt=conn.prepareStatement(strSql);
									mobilenum=getCellDataCust(rowCust.getCell(1));
									stmt.setString(1, getCellDataCust(rowCust.getCell(0)));
									stmt.setString(2, getCellDataCust(rowCust.getCell(1)));
									stmt.setString(3, getCellDataCust(rowCust.getCell(2)));
									ResultSet rs=stmt.executeQuery();
										
									//System.out.println("loop");
									while (rs.next()) {
									    //v_count=rs.getInt("v_count");
										
					                   /* worksheet_out.Cells[count, "B"].value = mobilenum;
			                            worksheet_out.Cells[count, "C"].value = accountnum;//
			                            worksheet_out.Cells[count, "D"].value = prodstatus;//prodstatus
			                            worksheet_out.Cells[count, "E"].value = productseqs;
			                            worksheet_out.Cells[count, "F"].value = tariffname;//
			                              worksheet_out.Cells[count, "G"].value = startdate;
			                             worksheet_out.Cells[count, "H"].value = enddate;*/
			                            
									        System.out.println(rs.getString("event_source"));
									        mobilenum=rs.getString("event_source");
						                    accountnum = rs.getString("account_num");
						                    prodstatus = rs.getString("product_status");
						                    productseqs = rs.getString("product_seq");
						                    tariffname = rs.getString("tariff_name");
				                            startdate = rs.getString("start_dat");
				                            enddate = rs.getString("end_dat");
				                            
				                            
				                            //08xxxxxxxx
				                            //Active, Suspend etc.
				                            //5GNSA, 5SA
				                            //4271
				                            //09/06/2022 00:00
				                            //09/06/2024 00:00
				                            
									}
									introw++;
									//System.out.println(introw);
									rs.close();
									// Create a row and put some cells in it. Rows are 0 based.
									row = sheetout.createRow((short)introw);
									// Create a cell and put a value in it.
									row.createCell(0).setCellValue(introw);
									row.createCell(1).setCellValue(mobilenum);
									row.createCell(2).setCellValue(accountnum);
									row.createCell(3).setCellValue(prodstatus);
									row.createCell(4).setCellValue(productseqs);
									row.createCell(5).setCellValue(tariffname);
									row.createCell(6).setCellValue(startdate);
									row.createCell(7).setCellValue(enddate);
									}//check account
							        

								    //System.out.println("out");
							}
							
							logger.info("new excel format ");
							workbook.write(fos);
					        fos.flush();
					        fos.close();
					        conn.close();
							
						} catch(OldExcelFormatException e){
							JOptionPane.showMessageDialog(null, "read excel error " +e);
					
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
						
							logger.info("old excel format ");
							
						}
						JOptionPane.showMessageDialog(null, "generate output file " + outputFile + " done");
					//}
					
					//}  for loop
					}// else

				//}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				//if (result!=JOptionPane.CANCEL_OPTION) {
			
					logger.error(e);
					//JOptionPane.showMessageDialog(null, "load file error");
					JOptionPane.showMessageDialog(null, "generate output file " + filename + "errr" +e);
					//JOptionPane.showMessageDialog(null, "old excel format" +e);
			           
				//}
			}
	}//end sub

	public static void processCheckPromotionFBB(String filenameinput,String filename,String extension,String pathInput,String username,String password,String database) {
		//String pathInput = "C:\\xlsFloder/";
		//System.out.println(pathInput);
		org.apache.poi.ss.usermodel.Sheet sheet;
		//org.apache.poi.ss.usermodel.Sheet sheetOut;
		//Sheet sheet;
		String strSql;
        String accountnum = null;
        String customerref = null;
        String tariffname = null;
        String tariffid=null;
        String productseqs = null;
        String  startdate = null;
        String enddate = null;
        String prodstatus = null;
        String mobilenum=null;
        String outputFile=null;
        String datenow =null;
		boolean v_check=true;
	    	  
			DbControl db = new DbControl();
			Connection conn = null;

			try {
				//if ((filename.length()>=0) ) {
		
					conn = db.getConnectionOra(username,password,database);
					DbServiceCheckProfileTools dbServicedbTools = new DbServiceCheckProfileTools(conn);
					if (conn !=null) {
					//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format"+conn);
						
					//dbService.deleteInvGenBatchAct();
					//logger.info("delete before insert");
				       
					//	if(("xlsx".equalsIgnoreCase(extension)) ||("xls".equalsIgnoreCase(extension)))  {
						//logger.info("read sheet :"+extension);
						logger.info("read file name :"+filenameinput);
						//JOptionPane.showMessageDialog(null, fileNameInput);
						//ArrayList<ExcelSheetHead> excelDataList = new ArrayList<ExcelSheetHead>();
						try {
							
							//JOptionPane.showMessageDialog(null, "new org.apache.poi.ss.usermodel.Workbook format");
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							
							//read file
							org.apache.poi.ss.usermodel.Workbook wb=WorkbookFactory.create(ExcelFileToRead);
							
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
							//JOptionPane.showMessageDialog(null, sheet.getLastRowNum());
							
							//output file
							
							
						    /*SimpleDateFormat formatter = new SimpleDateFormat("YYYYMMDD");  
						    Date date = new Date();  
						    System.out.println(formatter.format(date));  
						    */
							
						    long millis=System.currentTimeMillis();  
						      
						    // creating a new object of the class Date  
						    java.sql.Date date = new java.sql.Date(millis);       
						    ///System.out.println(date);   
						    
						    outputFile="Check_Promoton_FBB_Report_" + date + ".xlsx";
						    
							
							//XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
							//System.out.println("TEST");
							File fileName = new File(pathInput +outputFile);
						
					        FileOutputStream fos = new FileOutputStream(pathInput +outputFile);
					        XSSFWorkbook  workbook = new XSSFWorkbook();            

					        XSSFSheet sheetout = workbook.createSheet("Check_Promotion_FBB");  
					
				         /*     worksheet_out.Cells[1, "A"] = "No.";
			                    worksheet_out.Cells[1, "B"] = "FBB_ID";
			                    worksheet_out.Cells[1, "C"] = "BA No.";
			                    worksheet_out.Cells[1, "D"] = "Status FBB ID";
			                    worksheet_out.Cells[1, "E"] = "Product Seq. Promotion";
			                    worksheet_out.Cells[1, "F"] = "Billing Tariff ID"; 
			                    worksheet_out.Cells[1, "G"] = "Promotion_name";
			                    worksheet_out.Cells[1, "H"] = "Start date";
			                    worksheet_out.Cells[1, "I"] = "End date";*/
			                    
			                int introw=0;
			                Row row = sheetout.createRow((short)introw);   
							Cell cell = row.createCell(0);
							/*
						    FBB ID
							Account no.
							status FBB ID
							Product Seq. Promotion
							Billing Trariff ID
							Promotion name
							Start date
							End date*/
							cell.setCellValue("No.");
							row.createCell(1).setCellValue("FBB ID");
							row.createCell(2).setCellValue("BA No.");
							row.createCell(3).setCellValue("Status FBB ID");
							row.createCell(4).setCellValue("Product Seq. Promotion");
							row.createCell(5).setCellValue("Billing Tariff ID");
							row.createCell(6).setCellValue("Promotion Name");
							row.createCell(7).setCellValue("star_date");
							row.createCell(8).setCellValue("end_date");
							for (int i = 1; i <= lastRow; i++) {
								ExcelInputTemplate listRow = new ExcelInputTemplate();
								Row rowCust = sheet.getRow(i);
								//JOptionPane.showMessageDialog(null, rowCust.getCell(1));
								if((rowCust != null) && (rowCust.getCell(1).toString().length() >0)) {
									logger.info("row :"+(i-1));

									listRow.setmoBile(getCellDataCust(rowCust.getCell(1)));
		
									strSql= " select distinct (select ce1.event_source from custeventsource ce1 " +
					                                 " where ce1.customer_ref=ce.customer_ref  " +
					                                 " and ce1.event_type_id=1 " +
					                                 " and rownum<=1) event_source, e.account_num,c.customer_ref, d.product_seq,t.tariff_id ,t.tariff_name,to_char(d.START_DAT,'DD/MM/YYYY')start_dat, to_char(d.END_DAT,'DD/MM/YYYY') end_dat, " +
					                                 " (select decode(cpds.product_status,'OK','ACTIVE','SU','SUSPEND','TX','TERMINATE') from custproductstatus cpds " +
					                                 " where cpds.customer_ref=c.customer_ref and cpds.product_seq=d.product_seq and cpds.effective_dtm =(select  min(cpds1.effective_dtm) from custproductstatus cpds1 where cpds1.customer_ref=cpds.customer_ref and cpds1.product_seq=cpds.product_seq and cpds1.effective_dtm between add_months(acc.next_bill_dtm,-1) and acc.next_bill_dtm-1 )) product_status  " +
					                                 " from custproductdetails e,account acc,CUSTPRODUCTTARIFFDETAILS d,tariff  t,custeventsource ce, " +
					                                 " custhasproduct c " +
					                                 " where e.customer_ref=acc.customer_ref " +
					                                 " and e.account_num=acc.account_num " +
					                                 " and e.customer_ref=d.customer_Ref " +
					                                 " and  c.PRODUCT_SEQ=d.PRODUCT_SEQ " +
					                                 " and c.CUSTOMER_REF=d.CUSTOMER_REF " +
					                                 " and c.PRODUCT_SEQ=e.PRODUCT_SEQ " +
					                                 " and c.CUSTOMER_REF=e.CUSTOMER_REF " +
					                                 " and d.tariff_id=t.tariff_id " +
					                                 " and ce.event_source= ? "+
					                                 " and d.product_seq= ? " +
					                                 " and ce.customer_ref=e.customer_ref ";
									
									System.out.println(strSql);
									PreparedStatement stmt=conn.prepareStatement(strSql);
									mobilenum=getCellDataCust(rowCust.getCell(1));
									stmt.setString(1, getCellDataCust(rowCust.getCell(1)));
									stmt.setString(2, getCellDataCust(rowCust.getCell(2)));
									
									System.out.println(getCellDataCust(rowCust.getCell(1)));
									System.out.println(getCellDataCust(rowCust.getCell(2)));
									//stmt.setString(3, getCellDataCust(rowCust.getCell(1)));
									ResultSet rs=stmt.executeQuery();
										
									//System.out.println("loop");
									while (rs.next()) {
									    //v_count=rs.getInt("v_count");
										
					                   /*  worksheet_out.Cells[count, "A"] = count - 1;
			                            worksheet_out.Cells[count, "B"].value = mobilenum;
			                            worksheet_out.Cells[count, "C"].value = accountnum;//
			                            worksheet_out.Cells[count, "D"].value = prodstatus;//prodstatus
			                            worksheet_out.Cells[count, "E"].value = productseqs;
			                            worksheet_out.Cells[count, "F"].value = tariffid;//
			                            worksheet_out.Cells[count, "G"].value = tariffname;//
			                            worksheet_out.Cells[count, "H"].value = startdate;
			                            worksheet_out.Cells[count, "I"].value = enddate;*/
			                            
									        System.out.println(rs.getString("event_source"));
									        //mobilenum=getCellDataCust(rowCust.getCell(1));
						                    accountnum = rs.getString("account_num");
						                    prodstatus = rs.getString("product_status");
						                    productseqs = rs.getString("product_seq");
						                    tariffid=rs.getString("tariff_id");
						                    tariffname = rs.getString("tariff_name");
				                            startdate = rs.getString("start_dat");
				                            enddate = rs.getString("end_dat");
				                            
				                            
				                            //08xxxxxxxx
				                            //Active, Suspend etc.
				                            //5GNSA, 5SA
				                            //4271
				                            //09/06/2022 00:00
				                            //09/06/2024 00:00
				                            
									}
									introw++;
									//System.out.println(introw);
									rs.close();
									// Create a row and put some cells in it. Rows are 0 based.
									row = sheetout.createRow((short)introw);
									// Create a cell and put a value in it.
									row.createCell(0).setCellValue(introw);
									row.createCell(1).setCellValue(mobilenum);
									row.createCell(2).setCellValue(accountnum);
									row.createCell(3).setCellValue(prodstatus);
									row.createCell(4).setCellValue(productseqs);
									row.createCell(5).setCellValue(tariffid);
									row.createCell(6).setCellValue(tariffname);
									row.createCell(7).setCellValue(startdate);
									row.createCell(8).setCellValue(enddate);
									}//check account
							        

								    //System.out.println("out");
							}
							
							logger.info("new excel format ");
							workbook.write(fos);
					        fos.flush();
					        fos.close();
					        conn.close();
							
						} catch(OldExcelFormatException e){
							JOptionPane.showMessageDialog(null, "read excel error " +e);
					
							InputStream ExcelFileToRead = new FileInputStream(pathInput+filenameinput);
						
							HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);
		
							sheet = wb.getSheetAt(0);
							
							int lastRow = sheet.getLastRowNum();
						
							logger.info("old excel format ");
							
						}
						JOptionPane.showMessageDialog(null, "generate output file " + outputFile + " done");
					//}
					
					//}  for loop
					}// else

				//}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				//if (result!=JOptionPane.CANCEL_OPTION) {
			
					logger.error(e);
					//JOptionPane.showMessageDialog(null, "load file error");
					JOptionPane.showMessageDialog(null, "generate output file " + filename + "errr" +e);
					//JOptionPane.showMessageDialog(null, "old excel format" +e);
			           
				//}
			}
	}//end sub

	private static String getCellDataCust(Cell cell) throws Exception {
		String output = "";
		if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				output = cell.getStringCellValue();
			} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				if (cell.getNumericCellValue() % 1 == 0) {
					output = String.valueOf((int) cell.getNumericCellValue());
				} else {
					output = String.valueOf(cell.getNumericCellValue());
				}
			}
		}
		return output.trim();
	}
}
