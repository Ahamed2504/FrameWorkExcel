package com.interactwithexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadNumericAndAlpha {
	public static void main(String[] args) throws IOException {
		
		//Create Object for File
		File f = new File("D:\\Template for Program.xlsx");
		
		//To Read data
		FileInputStream stream = new FileInputStream(f);
		
		//To Create object for Workbook (Excel)
		Workbook w = new XSSFWorkbook(stream);
		
		//To get sheet from Workbook
		Sheet sheet = w.getSheet("Alpha & Numeric");
		
		//To Find how many no.of Rows filled with data
		//To Iterate or Separate each Row ------- Row
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			
		//To Iterate or Separate each Cell present in Excel ----- Column
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				
		//To determine which type of data stored in cell  ---- It will give either 1-String Or 0-Number
		int cellType = cell.getCellType();
		if (cellType==1) {
			String stringCellValue = cell.getStringCellValue();
			System.out.print(stringCellValue + "\t");
		}
		else {
			double numericCellValue = cell.getNumericCellValue();
			System.out.println();
			
			//Type Conversion It will convert into without decimal number
			long l = (long)numericCellValue;
			System.out.print(l + "\t");
		}
		
		}
			
			}
			
		}
	}

