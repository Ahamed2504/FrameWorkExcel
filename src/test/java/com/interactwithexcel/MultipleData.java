package com.interactwithexcel;         // [All data present in the Excel Sheet]

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MultipleData {
	public static void main(String[] args) throws IOException {
		
		//Create Object for File
		File f = new File("D:\\Template for Program.xlsx");
		
		//To Read data
		FileInputStream stream = new FileInputStream(f);
		
		//To Create object for Workbook (Excel)
		Workbook w = new XSSFWorkbook(stream);
		
		//To get sheet from Workbook
		Sheet sheet = w.getSheet("Alpha");
		
		//To Find how many no.of Rows filled with data
		//To Iterate or Separate each Row ------- Row
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			
		//To Iterate or Separate each Cell present in Excel ----- Column
		for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			Cell cell = row.getCell(j);
				
		//To get the data from the cells
		String stringCellValue = cell.getStringCellValue();
		System.out.print(stringCellValue+ "\t");	
			}
		System.out.println();
			
		}
	}

}
