package com.interactwithexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ModifyWriteData {
	
	public static void main(String[] args) throws IOException {
		
		// Create object for File
		File f = new File ("D:\\Dummy Excel\\Jan Excel.xlsx");
		
		//To Read the Data
		FileInputStream stream = new FileInputStream(f);
		
		//Create Object for Workbook
		Workbook w = new XSSFWorkbook(stream);
		
		//To get sheet from Workbook
		Sheet sheet = w.getSheet("Dummy");
				
		//To get the Row from the Sheet
		Row row = sheet.getRow(0);
				
		//To get the Cell from the Row
		Cell cell = row.getCell(0);
		
		//To get the data from the cell
		String stringCellValue = cell.getStringCellValue();
		if (stringCellValue.equalsIgnoreCase("shukkur ahamed a")) {
			cell.setCellValue("Mr & Ms Ahamed");
			
		}
		FileOutputStream streamOut = new FileOutputStream(f);
		w.write(streamOut);	
	}
			
}


