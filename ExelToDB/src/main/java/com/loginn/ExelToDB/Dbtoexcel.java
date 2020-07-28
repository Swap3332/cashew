package com.loginn.ExelToDB;

import java.io.File;
import java.io.FileOutputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Dbtoexcel {
   public static void main(String[] args) throws Exception {

      //Connecting to the database
      Class.forName("com.mysql.jdbc.Driver");
      Connection connect = DriverManager.getConnection("jdbc:mysql://localhost:3306/mydb1", "root" , "root345");

      //Getting data from the table emp_tbl
      Statement statement = connect.createStatement();
      ResultSet resultSet = statement.executeQuery("select * from stud");

      //Creating a Work Book
      XSSFWorkbook workbook = new XSSFWorkbook();

      //Creating a Spread Sheet
      XSSFSheet spreadsheet = workbook.createSheet("abc");
      XSSFRow row = spreadsheet.createRow(0);
      XSSFCell cell;
     
      cell = row.createCell(0);
      cell.setCellValue("Name");
     
      cell = row.createCell(1);
      cell.setCellValue("Age");
      /* 
      cell = row.createCell(3);
      cell.setCellValue("Height");
  
      cell = row.createCell(4);
      cell.setCellValue("SALARY");
     
      cell = row.createCell(5);
      cell.setCellValue("DEPT");
      */
      int i = 1;

      while(resultSet.next()) {
         row = spreadsheet.createRow(i);
         cell = row.createCell(0);
         cell.setCellValue(resultSet.getString("Name"));
         
         cell = row.createCell(1);
         cell.setCellValue(resultSet.getInt("Age"));
         /*    
         cell = row.createCell(3);
         cell.setCellValue(resultSet.getString("Height"));
         
        cell = row.createCell(4);
         cell.setCellValue(resultSet.getString("PERCENTAGE"));
         
         cell = row.createCell(5);
         cell.setCellValue(resultSet.getString("EMAIL"));  */
         i++;
      }
     
      FileOutputStream out = new FileOutputStream(
         new File("C:\\Users\\User\\Downloads\\xyz.xls"));
     
      workbook.write(out);
      out.close();
     
      System.out.println("xyz.xls written successfully");
   }
}
