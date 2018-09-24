package com.excel.example;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OpenExcelFile {

	   public static void main(String args[])throws Exception { 
	      File xclfile = new File("createworkbook.xlsx");   //"C:\\Users\\hmubaslat\\Desktop\\ex.xlsx"
	      FileInputStream fIP = new FileInputStream(xclfile);
	      
	      //Get the workbook instance for XLSX file 
	      XSSFWorkbook workbook = new XSSFWorkbook(fIP);
	     
	      if(xclfile.isFile() && xclfile.exists()) {
	         System.out.println("openworkbook.xlsx file open successfully.");
	      } else {
	         System.out.println("Error to open openworkbook.xlsx file.");
	      }
	   }
	}
	
