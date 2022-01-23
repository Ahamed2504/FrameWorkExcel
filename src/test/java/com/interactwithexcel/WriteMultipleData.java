package com.interactwithexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteMultipleData {

	public static void main(String[] args) throws IOException {
		
		// Create object for File
		File f = new File ("D:\\Dummy Excel\\Jan Excel.xlsx");
		
		//Create object for WorkBook
		Workbook w = new XSSFWorkbook();
		
		//Create Sheet in Excel
		Sheet createSheet = w.createSheet("Dummy");
		
		//Create a Multiple Rows in Excel by using Iterate
		for (int i = 0; i <=5; i++) {
			Row createRow = createSheet.createRow(i);
			
		//Create a Multiple Cells in Excel by using Iterate
		for (int j = 0; j <=5; j++) {
			Cell createCell = createRow.createCell(j);
			
		//Set Value in cell
		createCell.setCellValue("Thameer Ahamed S");
		createCell.setCellValue("Thameer Ahamed S");
		createCell.setCellValue("Thameer Ahamed S");
		createCell.setCellValue("Thameer Ahamed S");
		createCell.setCellValue("Thameer Ahamed S");
		
		//To Write a data in Excel
		FileOutputStream streamOut = new FileOutputStream(f);
		
		w.write(streamOut);
		
		}

	}
		
	}
}


