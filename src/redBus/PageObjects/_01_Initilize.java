package redBus.PageObjects;

import java.io.File;
import java.util.concurrent.TimeUnit;
import jxl.*;
import jxl.format.Colour;
import jxl.format.Pattern;
import jxl.write.Label;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class _01_Initilize extends UtilityScript {

 private WebDriver driver;

 public _01_Initilize(WebDriver driver) throws InterruptedException {
  this.driver = driver;
 }

 public _01_Initilize zOpen(String url) throws Exception {
  driver.manage().timeouts()
    .implicitlyWait(ImplicitWait, TimeUnit.SECONDS);
  // Code to mazimize the window. Reason some times Auto suggest,some
  // objects will fail if not maximized
  String script = "if (window.screen){window.moveTo(0,0);window.resizeTo(window.screen.availWidth,window.screen.availHeight);};";
  ((JavascriptExecutor) driver).executeScript(script);
  driver.get(url);
  return this;
  }
 public void Retake(String xPath, String value) throws Exception {
	 		try {
	 			//Thread.sleep(10000);
	 			driver.findElement(By.xpath(xPath)).click();
	 			} catch (Exception e) {
				System.out.println("Else Block");
				driver.findElement(By.xpath(value)).click(); //clicking on Retake 
				driver.findElement(By.xpath(xPath)).click(); //clicking on Start Assessment		
	 		} 
 }
 public void Groups()throws Exception {
	 String datavalue=null;
	 Workbook workbook = Workbook.getWorkbook(new File("C:\\MobileTesting\\TestScripts\\Premera Group Structure.xls"));
	 Sheet sheet = workbook.getSheet("Sheet3");
	 for(int row = 994; row < 13861; row ++)
	 {
		 Thread.sleep(8000);
		 driver.findElement(By.xpath("//table/tbody/tr/td/input[@id='btnAddGroup']")).click();
		 Thread.sleep(5000);
		 for(int col = 0; col < 3; col ++)
		 {
			switch(col)
			{
			 case 0:
				 datavalue =sheet.getCell(col,row).getContents();
				 System.out.println(datavalue);
				 driver.findElement(By.xpath("//div[@class='search-icon']/input[@id='groupsAddAutoComplete']")).sendKeys(datavalue);
				 Thread.sleep(4000);
				 driver.findElement(By.xpath("//div[@class='search-icon']/input[@id='groupsAddAutoComplete']")).click();
				 break;
			 case 1:
				 datavalue =sheet.getCell(col,row).getContents();
				 System.out.println(datavalue);
				 driver.findElement(By.xpath("//table/tbody/tr/td/input[@id='GroupName']")).sendKeys(datavalue);
				 break;
			 case 2:
				 datavalue =sheet.getCell(col,row).getContents();
				 System.out.println(datavalue);
				 driver.findElement(By.xpath("//table/tbody/tr/td/input[@id='grpCode']")).sendKeys(datavalue);
				 break;
			}
		}
		 driver.findElement(By.xpath("//table/tbody/tr/td/input[@id='btnSubmit']")).click();
		 Thread.sleep(15000);
		 xScreenShot();
	 }
 }
}