package com.virus;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Maleria {
public static void main(String[] args) throws EncryptedDocumentException, IOException {
	           FileInputStream fis = new FileInputStream(".\\src\\test\\resources\\newCity.xlsx");
	           Workbook wb = WorkbookFactory.create(fis);
	          Sheet ref = wb.getSheet("Sheet1");  
	          WebDriverManager.firefoxdriver().setup();
	          WebDriver driver=new FirefoxDriver();
	          driver.get("https://www.amazon.in/");
	          List<WebElement> ele = driver.findElements(By.xpath("//a"));
	          int count = ele.size();
	          for(int i=0;i<count;i++)
	          {
	        	  Row r = ref.createRow(i);
	        	  Cell c = r.createCell(0);
	        	  c.setCellValue(ele.get(i).getAttribute("href"));
	          }
	         // Row r = ref.getRow(6);
	          //Cell c = r.getCell(2);
	         //String Ans = c.toString();
	         //System.out.println(Ans);
	          
	          
	          //for(i=0;i<count;i++)
	        //  {
	        //	ref.getRow(0);
	        	
	        //  }
	  FileOutputStream wb2 = new FileOutputStream(".\\src\\test\\resources\\newCity.xlsx");
	        wb.write(wb2);
	        System.out.println(wb2);
	            
	
}
}
