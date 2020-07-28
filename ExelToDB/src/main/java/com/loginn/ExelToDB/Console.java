package com.loginn.ExelToDB;

import java.io.File;  
import java.io.FileInputStream;  
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
public class Console  
{  
public static void main(String args[]) throws IOException  
{  
	ArrayList data = new ArrayList();

//obtaining input bytes from a file  
FileInputStream fis=new FileInputStream(new File("C:\\Users\\User\\Downloads\\empdata.xlsx"));  
//creating workbook instance that refers to .xls file  
XSSFWorkbook wb=new XSSFWorkbook(fis);  
//creating a Sheet object to retrieve the object  
XSSFSheet sheet=wb.getSheetAt(0);  
//evaluating cell type  

Iterator itr = sheet.iterator();
while (itr.hasNext()) {
    Row rowitr = (Row) itr.next();
    Iterator cellitr = rowitr.cellIterator();
    while(cellitr.hasNext()) {
        Cell celldata = (Cell) cellitr.next();

        switch(celldata.getCellType()) {
        case STRING:
            data.add(celldata.getStringCellValue());
            System.out.print(celldata.getStringCellValue()+ "\t\t");  
            break;
        case NUMERIC:
            data.add(celldata.getNumericCellValue());
            System.out.print(celldata.getNumericCellValue()+ "\t\t");  
            break;
        case BOOLEAN:
            data.add(celldata.getBooleanCellValue());
            break;
        }
    }
}
/*
for (int i=0;i<data.size();i++) {
    if(data.get(i).equals("Sharan")) {
        System.out.println(data.get(i));
        System.out.println(data.get(i+1));
        System.out.println(data.get(i+2));
        System.out.println(data.get(i+3));          
    }
    if(data.get(i).equals("Kiran")) {
        System.out.println(data.get(i));
        System.out.println(data.get(i+1));
        System.out.println(data.get(i+2));
        System.out.println(data.get(i+3));
    }
    if(data.get(i).equals("Jhade")) {
        System.out.println(data.get(i));
        System.out.println(data.get(i+1));
        System.out.println(data.get(i+2));
        System.out.println(data.get(i+3));
}
*/
}       
}
