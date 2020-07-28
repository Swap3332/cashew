package com.loginn.ExelToDB;

import java.io.*;
import java.sql.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
     
public class Exceldb
{
      public static void main( String [] args )
      {
    String fileName="C:\\Users\\User\\Downloads\\empp.xls";
    Vector<Vector<HSSFCell>> dataHolder=read(fileName);
    saveToDatabase(dataHolder);
      }
      public static Vector<Vector<HSSFCell>> read(String fileName)  
      {
     Vector<Vector<HSSFCell>> cellVectorHolder = new Vector<Vector<HSSFCell>>();
     try
     {    
     FileInputStream myInput = new FileInputStream(fileName);
     POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
     HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
     HSSFSheet mySheet = myWorkBook.getSheetAt(0);
     Iterator<Row> rowIter = mySheet.rowIterator();
     while(rowIter.hasNext())
     {
     HSSFRow myRow = (HSSFRow) rowIter.next();
     Iterator<Cell> cellIter = myRow.cellIterator();
     Vector<HSSFCell> cellStoreVector=new Vector<HSSFCell>();
     while(cellIter.hasNext())
     {
      HSSFCell myCell = (HSSFCell) cellIter.next();
      cellStoreVector.addElement(myCell);
     }
     cellVectorHolder.addElement(cellStoreVector);
     }
     }
     catch (Exception e)
     {
     e.printStackTrace();
     }
     return cellVectorHolder;
      }
      private static void saveToDatabase(Vector<Vector<HSSFCell>> dataHolder)
      {
     String Name="";
     String Age="";
     for (int i=0;i<dataHolder.size(); i++)
     {
     Vector<HSSFCell> cellStoreVector=dataHolder.elementAt(i);
     for (int j=0; j < cellStoreVector.size();j++)
     {
     HSSFCell myCell = (HSSFCell)cellStoreVector.elementAt(j);
     String st = myCell.toString();
     Name=st.substring(0,2);
     Age=st.substring(0);
     }
     try
     {
     Class.forName("com.mysql.jdbc.Driver").newInstance();
     Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/mydb1","root", "root345");
     Statement stat=con.createStatement();
     stat.executeUpdate("insert into stud(Name,Age) value('"+Name+"','"+Age+"')");
     System.out.println("Data is inserted");
     stat.close();
     con.close();
     }
     catch(Exception e){}
     }
}
}       