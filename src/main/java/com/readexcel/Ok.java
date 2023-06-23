package com.readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Ok {

	public static void main(String[] args) {
		
		try  
		{  
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection con=DriverManager.getConnection( "jdbc:mysql://localhost:3306/jaydb","root","J@ypatidar200503");
			
			File file = new File("L:\\boats.xlsx");  
			FileInputStream fis = new FileInputStream(file);    
		
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);      

			PreparedStatement pstm;  
			pstm = con.prepareStatement("truncate table new_table");
			pstm.execute();

			Row row;
			for(int i=1; i<=sheet.getLastRowNum(); i++){	
        	
				row = sheet.getRow(i);
				
				String fn = row.getCell(0).getStringCellValue();
				String ln = row.getCell(1).getStringCellValue();	
				String addr = row.getCell(2).getStringCellValue();
            
	            pstm = con.prepareStatement("insert into new_table values(?,?,?)");
	            pstm.setString(1, fn);
	            pstm.setString(2, ln);
	            pstm.setString(3, addr);

	            pstm.execute();
            
        }
        con.close();
        System.out.println("Successfully transfered excel data to mysql table");
		  
		}catch(ClassNotFoundException e){
            System.out.println(e);
        }
		catch(IOException ioe){
            System.out.println(ioe);
		}catch(Exception ex){
            System.out.println(ex);
        }
	}

}
