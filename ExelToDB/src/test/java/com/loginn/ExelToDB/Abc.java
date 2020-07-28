package com.loginn.ExelToDB;

import java.io.*;
import java.util.Iterator;
import java.sql.*;
import java.util.*;
import java.util.Date;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Abc {
	public static void main(String[] args) {
        String jdbcURL = "jdbc:mysql://localhost:3306/mydb1";
        String username = "root";
        String password = "root345";
 
        String excelFilePath = "C:/Users/User/Downloads/xyz.xls";
 
        int batchSize = 20;
 
        Connection connection = null;
 
        try {
            long start = System.currentTimeMillis();
             
            FileInputStream inputStream = new FileInputStream(excelFilePath);
 
            Workbook workbook = new XSSFWorkbook(inputStream);
 
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
 
            connection = DriverManager.getConnection(jdbcURL, username, password);
            connection.setAutoCommit(false);
  
            String sql = "INSERT INTO stud (name, age) VALUES (?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);    
             
            int count = 0;
          
            //Iterator itr = firstSheet.iterator();
            rowIterator.next(); // skip the header row
            
            while (rowIterator.hasNext()) {
            
            Row rowitr = (Row) rowIterator.next();
            Iterator cellitr = rowitr.cellIterator();
            while(cellitr.hasNext()) {
                Cell celldata = (Cell) cellitr.next();
                int columnIndex =  celldata.getColumnIndex();
                int rowIndex = celldata.getRowIndex();
                
                switch(columnIndex) {
                case 0:
                	String name = celldata.getStringCellValue();
                	statement.setString(1, name);
                    System.out.print(celldata.getStringCellValue()+ "\t\t");  
                    break;
                case 1:
                	int age = (int) celldata.getNumericCellValue();
                	statement.setInt(2, age);
                    System.out.print(celldata.getNumericCellValue()+ "\t\t");  
                    break;
                case 2:
                    break;
                }
            }
        
            
           /*
           while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
 
                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();
 
                    int columnIndex =  nextCell.getColumnIndex();
 
                    switch (columnIndex) {
                    case 0:
                        String name = nextCell.getStringCellValue();
                        statement.setString(1, name);
                        break;
                    case 1:
                        Date enrollDate = nextCell.getDateCellValue();
                        statement.setTimestamp(2, new Timestamp(enrollDate.getTime()));
                    case 2:
                        int progress = (int) nextCell.getNumericCellValue();
                        statement.setInt(3, progress);
                    }
 
                }
             */    
                statement.addBatch();
                 
                if (count % batchSize == 0) {
                    statement.executeBatch();
                }              
 
            }
 
            workbook.close();
             
            // execute the remaining queries
            statement.executeBatch();
  
            connection.commit();
            connection.close();
             
            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));
             
        } catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }
 
    }
}

