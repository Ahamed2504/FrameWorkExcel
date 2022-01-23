package com.interactwithexcel;        // [To Read particular data from Excel]

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ReadExcelData {
	public static void main(String[] args) throws IOException {
		
		//Create Object for File
		File f = new File ("D:\\Template for Program.xlsx");
		
		//To Read data
		FileInputStream stream = new FileInputStream(f);
		
		//To Create object for Workbook (Excel)
		Workbook w = new XSSFWorkbook(stream);
		
		//To get sheet from Workbook
		Sheet sheet = w.getSheet("Alpha");
		
		//To get the Row from the Sheet
		Row row = sheet.getRow(1);
		
		//To get the Cell from the Row
		Cell cell = row.getCell(2);
		
		//To get the data from the cell
		String stringCellValue = cell.getStringCellValue();
		System.out.println(stringCellValue);		
}
}
