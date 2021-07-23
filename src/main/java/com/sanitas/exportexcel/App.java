package com.sanitas.exportexcel;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import oracle.jdbc.OracleTypes;
import org.apache.poi.ss.usermodel.*;
 
import java.util.*;
import java.io.*;
import java.sql.*;
import java.text.*;

import org.apache.log4j.PropertyConfigurator;
import org.apache.log4j.Logger;

public class App 
{
	private static final Logger logger = Logger.getLogger(App.class);
	
	//Variables a usar desde el archivo properties.
	static String DB_IP = PropertiesHelper.getParameter("DB_IP");
	static String DB_Port = PropertiesHelper.getParameter("DB_Port");
	static String DB_Instance = PropertiesHelper.getParameter("DB_Instance");
	static String DB_Username = PropertiesHelper.getParameter("DB_Username");
	static String DB_Password = PropertiesHelper.getParameter("DB_Password");
	static String DB_Stored_Procedure = PropertiesHelper.getParameter("DB_Stored_Procedure");
	static String DB_SP_Parameter = PropertiesHelper.getParameter("DB_SP_Parameter");
	static String Path_to_File = PropertiesHelper.getParameter("Path_to_File");
	static String File_Name = PropertiesHelper.getParameter("File_Name");
	static String Date_Format = PropertiesHelper.getParameter("Date_Format");
	static String Hour_Format = PropertiesHelper.getParameter("Hour_Format");
	
	public static void main( String[] args ) throws Exception
    {
		PropertyConfigurator.configure("log4j.properties");
		
		try {
			  logger.info("Proceso iniciado.");
			  SimpleDateFormat fecha_hoy = new SimpleDateFormat(Date_Format);
			  SimpleDateFormat hora_hoy = new SimpleDateFormat(Hour_Format);
			  
		      Calendar c = Calendar.getInstance();
		      c.setTime(new java.util.Date());
		      c.add(Calendar.DATE, 0);
		      String output = fecha_hoy.format(c.getTime());
		      String output_2 = hora_hoy.format(c.getTime());
		      
		      Connection con = getOracleConnection();
		      CallableStatement callableStmt = null;
		      ResultSet rst = null;
			  callableStmt = con.prepareCall("{ call "+ DB_Stored_Procedure +" }");
			  
			  callableStmt.setString(1, DB_SP_Parameter);// set the INPUT_NUM parameter
			  callableStmt.registerOutParameter(2, OracleTypes.CURSOR);// register the type of the out parameter for CURSOR
			  callableStmt.execute(); //execute stored procedure
			  rst = (ResultSet) callableStmt.getObject(2);//get cursor output
			  
		      /*
				XSSFWorkbook	-> Utilizado inicialmente pero procaba error "Java heap space" (Colapso de memoria de la JVM) 
				ya que da acceso a todas las filas del documento.
				
				SXSSFWorkbook	-> Está diseñado para vaciar filas en el disco en lugar de mantenerlas en la memoria. 
				Las filas antiguas que ya no se encuentran en la ventana se vuelven inaccesibles, ya que se escriben en el disco.
				Se trata de una clase mejorada de "XSSFWorkbook" utilizada en este proyecto.
		      */
			  
			  Workbook workbook = new SXSSFWorkbook();
			  Sheet  sheet = workbook.createSheet();
		      CreationHelper createHelper = workbook.getCreationHelper();
		      short dateFormat = createHelper.createDataFormat().getFormat("dd/MM/yyyy");
		      
		      //ESTILO_TITULO
		      XSSFFont font_title = (XSSFFont)workbook.createFont();
		      font_title.setBold(true);
		      XSSFCellStyle style_title = (XSSFCellStyle)workbook.createCellStyle();
		      style_title.setFont(font_title);
		      
		      //ESTILO_CUERPO
		      XSSFFont font_body = (XSSFFont)workbook.createFont();
		      font_body.setBold(false);
		      XSSFCellStyle style_body = (XSSFCellStyle)workbook.createCellStyle();
		      style_body.setFont(font_body);
		      
		      //ESTILO_COLUMNA_FECHA
		      XSSFFont font_date = (XSSFFont)workbook.createFont();
		      XSSFCellStyle style_date = (XSSFCellStyle)workbook.createCellStyle();
		      style_date.setFont(font_date);
		      style_date.setDataFormat(dateFormat);
		      
		      //CABECERA DE ARCHIVO
		      int rownum = 0;
		      int cellnum = 0;
		      ResultSetMetaData rsmd = rst.getMetaData();
		      Row row = sheet.createRow(rownum++);
		      for (int col = 1; col <= rsmd.getColumnCount(); col++) {
		    	  Cell cell = row.createCell(cellnum++);
		    	  cell.setCellValue((String)rsmd.getColumnName(col));
		    	  cell.setCellStyle(style_title);
		      }
		      
		      //CUERPO DE ARCHIVO
		      while (rst.next()) {
		        row = sheet.createRow(rownum++);
		        cellnum = 0;
		        		        
		        for (int col = 1; col <= rsmd.getColumnCount(); col++) {
		          
		        	Cell cell = row.createCell(cellnum++);
		        	if (rsmd.getColumnTypeName(col).equals("VARCHAR2"))
		        		cell.setCellValue(rst.getString(col));
		        		cell.setCellStyle(style_body);
		        	
		        	if (rsmd.getColumnTypeName(col).equals("DATE"))
		        		cell.setCellStyle(style_date);
		        		cell.setCellValue(rst.getString(col));
		        				        
		        	if (rsmd.getColumnTypeName(col).equals("NUMBER"))
			            cell.setCellValue(rst.getDouble(col));
		        		cell.setCellStyle(style_body);
		         }
		      }
		      try {
		    	  FileOutputStream out = new FileOutputStream(new File(Path_to_File + File_Name + output + "_" + output_2 + ".xlsx"));
		    	  workbook.write(out);
		    	  workbook.close();
		    	  logger.info("Proceso finalizado.");
		      }
		      catch (Exception e) {
		    	  e.printStackTrace();
		    	  logger.error(e);
		      }
		    }
		    catch (Exception e) {
	    		  e.printStackTrace();
	    		  logger.error(e);
		    }
    }
	
	public static Connection getOracleConnection() throws Exception {
		String driver = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@" + DB_IP +":" + DB_Port + ":" + DB_Instance;
		String username = DB_Username;
		String password = DB_Password;

		Class.forName(driver); // load Oracle driver
	    Connection conn = DriverManager.getConnection(url, username, password);
	    return conn;
	  }
}