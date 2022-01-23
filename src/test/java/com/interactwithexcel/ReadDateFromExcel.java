package com.interactwithexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDateFromExcel {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File("D:\\Template for Program.xlsx");
		
		FileInputStream stream = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(stream);
		
		Sheet sheet = w.getSheet("Combination");
		
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				
				int cellType = cell.getCellType();
				
				if (cellType==1) {
					String stringCellValue = cell.getStringCellValue();
					System.out.print(stringCellValue + "\t");
					
				}
				else if(DateUtil.isCellDateFormatted(cell)) {
					Date dateCellValue = cell.getDateCellValue();
					System.out.println(dateCellValue);
					
					//To Want desired Date Format
					SimpleDateFormat s = new SimpleDateFormat("MMM/yy/dd");
					String format = s.format(dateCellValue);
					System.out.println(format);
				}
					else {
						System.out.println();
						double numericCellValue = cell.getNumericCellValue();
						
						//Type Conversion
						long l = (long)numericCellValue;
						System.out.print(l + "\t");
					}
					
				}
					
				}
			
	}
}
