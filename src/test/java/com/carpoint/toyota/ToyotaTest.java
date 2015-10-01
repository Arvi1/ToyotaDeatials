package com.carpoint.toyota;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;

public class ToyotaTest {
	WebDriver driver;
	FileInputStream fis;
	HSSFWorkbook wb;
	HSSFSheet sheet;
	DateFormat dateFormat;
	Date date;
	
	  @BeforeTest
	  public void beforeTest() throws Exception {
		  driver = new FirefoxDriver();
		  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		  driver.manage().window().maximize();
		  
		  driver.get("http://www.carpoint.com.au/");
		  
		  dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		  date = new Date();
		  
		  fis = new FileInputStream(new File("D:/EclipseProjects/SeekProject/CarPointDemoProject/CarPoint.xls"));
		  wb = new HSSFWorkbook(fis);
		  sheet = wb.getSheet("Sheet1");
		
	  }
	  
	  @Test
	  public void selectToyota() throws Exception {
		  
		// Select make from the dropdown
		  WebElement drpdwnMake = driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_ddlMake"));
			List<WebElement> makeOptions = drpdwnMake.findElements(By.tagName("option"));
			for (WebElement option : makeOptions ) {
				if(option.getText().contains("Toyota"))
					option.click();
			}
			Thread.sleep(2000);
		// Select model from the dropdown
		  WebElement drpdwnModel = driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_ddlModel"));
			List<WebElement> modelOptions = drpdwnModel.findElements(By.tagName("option"));
			for (WebElement option : modelOptions ) {
				if(option.getText().contains("Corolla"))
					option.click();
			}
			Thread.sleep(2000);	
			
		// Select Body Type from the dropdown
		  WebElement drpdwnBodyType = driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_ddlBodyType"));
			List<WebElement> bodyTypeOptions = drpdwnBodyType.findElements(By.tagName("option"));
			for (WebElement option : bodyTypeOptions ) {
				if(option.getText().contains("Sedan"))
					option.click();
			}
			Thread.sleep(2000);	
	  // Select State from the dropdown
		  WebElement drpdwnState = driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_ddlState"));
			List<WebElement> stateOptions = drpdwnState.findElements(By.tagName("option"));
			for (WebElement option : stateOptions ) {
				if(option.getText().contains("Queensland"))
					option.click();
			}
			Thread.sleep(2000);	
	 // Select State from the dropdown
		  WebElement drpdwnRegion = driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_ddlRegion"));
			List<WebElement> regionOptions = drpdwnRegion.findElements(By.tagName("option"));
			for (WebElement option : regionOptions ) {
				if(option.getText().contains("Brisbane All"))
					option.click();
			}
			Thread.sleep(2000);	
	 // Select State from the dropdown
		  WebElement drpdwnPriceMax = driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_ddlPriceTo"));
			List<WebElement> priceMaxOptions = drpdwnPriceMax.findElements(By.tagName("option"));
			for (WebElement option : priceMaxOptions ) {
				if(option.getText().contains("5000"))
					option.click();
			}
			Thread.sleep(2000);
			
			driver.findElement(By.id("ctl07_p_d_ctl05_ctl01_ctl03_ctl01_btnSubmit")).click();
			
			Thread.sleep(4000);
			
		// Select Sort type from the dropdown	
			driver.findElement(By.xpath("//*[@id='csn-select-ctl09_p_ctl02_ctl04_sortControl']/div[1]/span/i")).click();
			Thread.sleep(2000);					  
			driver.findElement(By.linkText("Last updated")).click();
			Thread.sleep(4000);
			
		//  xPaths for car details	
			String cd1 = "//html/body/form/div[5]/div[2]/div[2]/div[2]/div[2]/div[2]/div[";
			String cd2 = "]/h2/a";
			
		//  xPaths for car price			
			String p1 = "//html/body/form/div[5]/div[2]/div[2]/div[2]/div[2]/div[2]/div[";
			String p2 = "]/div/div[2]/div[1]/p/a";

		//  for loop implemented to get price and car details  
			try {
			
				for (int i = 1; i < 6; i++){
					Cell resDate = sheet.getRow(i).getCell(1);
					Cell resCarDetails = sheet.getRow(i).getCell(2);
					Cell rescarPrice = sheet.getRow(i).getCell(3);
					String carDetails = driver.findElement(By.xpath(cd1+i+cd2)).getText().toString();				
					String carPrice = driver.findElement(By.xpath(p1+i+p2)).getText().toString();
					resCarDetails.setCellValue(carDetails);
					rescarPrice.setCellValue(carPrice);
					resDate.setCellValue(dateFormat.format(date));
					System.out.println(i+ " -- " + carDetails);
					System.out.println(i+ " -- " + carPrice);
				}
				Thread.sleep(2000);
				for (int i = 7; i < 9; i++){
					Cell resDate = sheet.getRow(i).getCell(1);
					Cell resCarDetails = sheet.getRow(i).getCell(2);
					Cell rescarPrice = sheet.getRow(i).getCell(3);
					String carDetails = driver.findElement(By.xpath(cd1+i+cd2)).getText().toString();				
					String carPrice = driver.findElement(By.xpath(p1+i+p2)).getText().toString();
					resCarDetails.setCellValue(carDetails);
					rescarPrice.setCellValue(carPrice);
					resDate.setCellValue(dateFormat.format(date));
					System.out.println(i+ " -- " + carDetails);
					System.out.println(i+ " -- " + carPrice);
				}
				Thread.sleep(2000);
				for (int i = 10; i < 12; i++){
					Cell resDate = sheet.getRow(i).getCell(1);
					Cell resCarDetails = sheet.getRow(i).getCell(2);
					Cell rescarPrice = sheet.getRow(i).getCell(3);
					String carDetails = driver.findElement(By.xpath(cd1+i+cd2)).getText().toString();				
					String carPrice = driver.findElement(By.xpath(p1+i+p2)).getText().toString();
					resCarDetails.setCellValue(carDetails);
					rescarPrice.setCellValue(carPrice);
					resDate.setCellValue(dateFormat.format(date));
					System.out.println(i+ " -- " + carDetails);
					System.out.println(i+ " -- " + carPrice);
				}
				Thread.sleep(2000);
				wb.close();
				fis.close();
				
				FileOutputStream fos = new FileOutputStream(new File("D:/EclipseProjects/SeekProject/CarPointDemoProject/CarPoint-Result4.xls"));
				wb.write(fos);
				fos.close();
			} catch (Exception e) {
				e.getMessage();
			}
			
	  }
	  
	  @AfterTest
	  public void afterTest() {
		  driver.quit();
	  }

}
