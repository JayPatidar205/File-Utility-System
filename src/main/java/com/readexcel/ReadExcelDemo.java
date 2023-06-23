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

public class ReadExcelDemo {

	public static void main(String[] args) {
		
		try  
		{  
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection con=DriverManager.getConnection( "jdbc:mysql://localhost:3306/jaydb","root","J@ypatidar200503");
			
			File file = new File("J:\\Book.xlsx");  
			FileInputStream fis = new FileInputStream(file);    
		
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);      

			PreparedStatement pstm;  
			pstm = con.prepareStatement("truncate table employee");
			pstm.execute();

			Row row;
			for(int i=1; i<=sheet.getLastRowNum(); i++){	
        	
				row = sheet.getRow(i);
				int id = (int) row.getCell(0).getNumericCellValue();
				String name = row.getCell(1).getStringCellValue();
				int salary= (int) row.getCell(2).getNumericCellValue();
				String department = row.getCell(3).getStringCellValue();
            
	            pstm = con.prepareStatement("insert into employee values(?,?,?,?)");
	            pstm.setInt(1, id);
	            pstm.setString(2, name);
	            pstm.setInt(3, salary);
	            pstm.setString(4, department);

	            pstm.execute();
            
        }
        con.close();
        System.out.println("Successfully transfered excel data to mysql table");
		  
		}
			
//			Iterator<Row> itr = sheet.iterator();
//			pstm = con.prepareStatement("insert into employee(ID, Name, Salary,Department) values(?,?,?,?)");
//
//			itr.next();
//			while (itr.hasNext())                 
//			{  
//			Row row = itr.next();  
//			Iterator<Cell> cellIterator = row.cellIterator();
//			
//			while (cellIterator.hasNext())   
//			{  
//				
//			Cell cell = cellIterator.next(); 
//			int columnIndex = cell.getColumnIndex();
//			
//			switch (columnIndex)               
//			{  
//		
//			case 0:    
//				int id= (int) cell.getNumericCellValue();
//				pstm.setInt(1, id);
//				pstm.execute();
//				break;
//				  
//			case 1:
//				String name= cell.getStringCellValue();
//				pstm.setString(2, name);
//				pstm.execute();
//				break;
//				 
//				
//			case 2:
//				int salary=(int) cell.getNumericCellValue();
//				pstm.setInt(3, salary);
//				pstm.execute();
//				break;
//				
//			case 3:
//				String department=cell.getStringCellValue();
//				pstm.setString(4, department);
//				pstm.execute();
//			 
//			 
//			}  
//			}  
//			
//			}  
			
        
       catch(ClassNotFoundException e){
            System.out.println(e);
        }
		catch(IOException ioe){
            System.out.println(ioe);
		}catch(Exception ex){
            System.out.println(ex);
        }
        }

}


