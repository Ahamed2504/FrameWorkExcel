package com.Framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.SendKeysAction;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDrivenTask {
	
	public static void main(String[] args) throws IOException {
		
		//Create object for File
		File f = new File("D:\\Dummy Excel\\Book1.xlsx");
		
		//To Read data
		FileInputStream stream = new FileInputStream(f);
		
		//Create object for WorkBook(Excel)
		Workbook w = new XSSFWorkbook(stream);
		
		//To get the Row from the Sheet
		Sheet sheet = w.getSheet("Sheet1");
		
		//To Set the Property
		WebDriverManager.chromedriver().setup();
		
		//Initialize WebDriver
		WebDriver driver = new ChromeDriver();
		
		//To Launch the URL
		driver.get("https://demoqa.com/forms");
		driver.manage().window().maximize();
		
		//Create a object for Actions class
		Actions a = new Actions(driver);
		
		//To Inspect Register Page by using XPath
		WebElement form = driver.findElement(By.xpath("//span[contains(text(),'Practice Form')]"));
		a.moveToElement(form).perform();
		a.click().perform();
		
		//To inspect WebElement by using Locators
		WebElement firstName = driver.findElement(By.id("firstName"));
		firstName.sendKeys(sheet.getRow(0).getCell(0).getStringCellValue());
		
		WebElement lastName = driver.findElement(By.id("lastName"));
		lastName.sendKeys(sheet.getRow(0).getCell(1).getStringCellValue());
		
		WebElement emailId = driver.findElement(By.id("userEmail"));
		emailId.sendKeys(sheet.getRow(0).getCell(2).getStringCellValue());
		
		JavascriptExecutor js = (JavascriptExecutor)driver;
		
		WebElement gender = driver.findElement(By.xpath("(//input[@name='gender'])[1]"));
		js.executeScript("arguments[0].click", gender);
		
		WebElement mobileNo = driver.findElement(By.id("userNumber"));
		mobileNo.sendKeys(sheet.getRow(0).getCell(3).getStringCellValue());
		
		WebElement sub = driver.findElement(By.id("subjectsContainer"));
			
			
			
		}
		}
	


