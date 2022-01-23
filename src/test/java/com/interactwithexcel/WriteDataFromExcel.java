package com.interactwithexcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataFromExcel {
	
	public static void main(String[] args) throws IOException {
		
	// Create object for File
	File f = new File ("D:\\Dummy Excel\\Jan Excel.xlsx");
	
	//Create object for WorkBook
	Workbook w = new XSSFWorkbook();
	
	//Create Sheet in Excel
	Sheet createSheet = w.createSheet("Dummy");
	
	//Create a Row in Excel
	Row createRow = createSheet.createRow(0);
	
	//Create a Cell in Excel
	Cell createCell = createRow.createCell(0);
	
	//Set Value in cell
	createCell.setCellValue("Thameer Ahamed S");
	
	//To Write a data in Excel
	FileOutputStream streamOut = new FileOutputStream(f);
	
	w.write(streamOut);
	
	}

}
