package redBus;

import org.openqa.selenium.By;
import org.openqa.selenium.By.ByName;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.*;
import org.openqa.selenium.ie.*;
import org.openqa.selenium.android.AndroidDriver;
import org.openqa.selenium.chrome.*;
import org.openqa.selenium.safari.*;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.*;
import org.openqa.selenium.support.ui.Select;
import org.seleniumhq.jetty7.util.log.Log;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.TimeUnit;
import java.util.List;
import jxl.*;
import jxl.format.Colour;
import jxl.format.Pattern;
import jxl.write.Label;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import static java.nio.file.StandardCopyOption.*;
	
	import redBus.PageObjects.UtilityScript;
import redBus.PageObjects._01_Initilize;

	public class RedBus extends UtilityScript 
	{
	public static WebDriver driver;
				
	 @BeforeClass(alwaysRun = true)
	 protected void setUp() throws Exception {
		 xKillExcel();
		 String browser =xBrowser();
		 	
		 
		 if (browser.equals("Android"))
		 {
			 driver = new AndroidDriver();
		 }
		 /*
		  
		   if (browser.equals("Firefox"))
		 {driver = new FirefoxDriver();}
		 if (browser.equals("IE"))
		 {System.setProperty("webdriver.ie.driver", "C:\\IEDriverServer.exe");
		 driver = new InternetExplorerDriver();}
		 if (browser.equals("Chrome"))
		 {System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
		 driver = new ChromeDriver();}
		 if (browser.equals("Safari"))
		 {driver = new SafariDriver();}
		 
		 */
		// driver.manage().window().maximize();
	  	// driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	  	//driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		 }
	 
	 @AfterClass(alwaysRun = true)
	 protected void tearDown() throws Exception {
	  driver.quit();
	  xKillIEs();
	 }

	 @Test(groups = {"RedBusTestCases"}, enabled = true)
	
	 public void Test() throws Exception,InterruptedException {
	  
	  _01_Initilize Initilize = PageFactory.initElements(driver,_01_Initilize.class);
	  String TestSuiteName = xTestSuite();
	  MethodName = MethodName + TestSuiteName; //Used while sending email to report the test cases executed
	  Method = TestSuiteName; //Used while sending email to report the test cases executed
	  Print("Start:" + xGetDateTimeIP());
	  String DataTime = xGetDateTime();
	  String TestPath = xTestPath();
	  String ResultPath = xTestResult();
	  File fileExisting = new File(ResultPath+"\\"+TestSuiteName+"_Result.xls");  
	  if (fileExisting.exists()){  
		 	  Path RSource = Paths.get(ResultPath+"\\"+TestSuiteName+"_Result.xls");
		 	  Path RTarget= Paths.get(ResultPath+"\\"+TestSuiteName+"_Result_"+DataTime+".xls");
		 	  Files.copy(RSource, RTarget);
	  }
	  Path source = Paths.get(TestPath);
	  Path target = Paths.get(ResultPath+"\\"+TestSuiteName+"_Result.xls");
	  Files.copy(source, target,REPLACE_EXISTING);
	  Workbook workbook = Workbook.getWorkbook(new File(TestPath));
	  WritableWorkbook Writebook = Workbook.createWorkbook(new File(ResultPath+"\\"+TestSuiteName+"_Result.xls"),workbook);
	  WritableSheet WriteSheet = Writebook.createSheet("Results",2); 
	  Sheet sheet = workbook.getSheet("TestScript");
	  int startRow, startCol, endRow, endCol,ci,cj, Rcol, RowNo;
	  String  temp = null,temp2 = null, property = null, Action = null, StepNo = null,expectedResult = null,value = null,xPath = null,Error =null,Function =null, DataSetValue=null, Handle = null, HandleBefore=null; 
	  Cell tableStart=sheet.findCell("TestScriptStart");
	  startRow=tableStart.getRow();
	  startCol=tableStart.getColumn();
	  Cell tableEnd= sheet.findCell("TestScriptEnd", startCol,startRow, 100, 64000,  false);

		         endRow=tableEnd.getRow();
		         endCol=tableEnd.getColumn();
		         ci=0;
		         WriteSheet.addCell(new Label(0,0,"DataSetNo",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(1,0,"StepNo",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(2,0,"PageName",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(3,0,"Property",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(4,0,"FieldName",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(5,0,"xPath",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(6,0,"Value",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(7,0,"ExpectedResult",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(8,0,"Action",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(9,0,"ActualResult",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(10,0,"Result",xFillCell(Colour.GRAY_50)));
		         WriteSheet.addCell(new Label(11,0,"Error",xFillCell(Colour.GRAY_50)));
		         RowNo=1;
		         for (int dcol=16;dcol<250;dcol++)
		         {
		        	 try {DataSetValue =sheet.getCell(dcol,1).getContents();} 
		        	 catch (Exception e) {break;}
		        	 String DataSkip = sheet.getCell(dcol,2).getContents();
		        	 DataSkip = DataSkip.toUpperCase();
		         if(DataSkip.equals("SKIP_ALL")){continue;} 
		         if(DataSetValue.equals(null)||DataSetValue == ""){break;}
		         else
		         {
		        	 TestCaseLoop:
		        	 for (int i=startRow+1;i<endRow;i++,ci++){
		             cj=0; Rcol =0;
		             StepNo=sheet.getCell(endCol-8,i).getContents();
		             if(StepNo.contains("?")==false){
		             for (int j=startCol+1;j<endCol;j++,cj++,Rcol++){
		            	 if (Rcol==0){WriteSheet.addCell(new Label(Rcol,RowNo,DataSetValue)); j = j-1;}
		            	 else if(Rcol==6){temp =sheet.getCell(dcol,i).getContents();
		            	 	WriteSheet.addCell(new Label(Rcol,RowNo,temp));j = j-1;}
		            	 else {temp2 =sheet.getCell(j,i).getContents();
		            	 	WriteSheet.addCell(new Label(Rcol,RowNo,temp2));}
		             }
		             RowNo++;
		             }
				//StepNo=sheet.getCell(endCol-8,i).getContents();
				property=sheet.getCell(endCol-5,i).getContents();
				xPath=sheet.getCell(endCol-4,i).getContents();
				value=sheet.getCell(dcol,i).getContents();
				expectedResult=sheet.getCell(endCol-3,i).getContents();
				Action=sheet.getCell(endCol-2,i).getContents();
				Function = sheet.getCell(endCol-1,i).getContents();
				property = property.toUpperCase();
				String TempValue = value.toUpperCase();
				Action = Action.toUpperCase();
					if((StepNo.contains("?"))||(TempValue.equals("SKIP"))){Action = "SKIP";}						
						
						switch(Action){
							
							case "OPEN":
									if(property.equals("PAGETITLE")){try {
										driver.get(value);
										HandleBefore = driver.getWindowHandle();} 
									catch (Exception e) { 
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
									break;
							case "TYPE":
									if(property.equals("TEXTFIELD")){try {
										//driver.findElement(By.name(xPath)).sendKeys(value); 
										driver.findElement(By.id(xPath)).sendKeys(value);}
										catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}			
									break;
							case "CLICK":
								 	if(property.equals("LINK")){try {
								 		//driver.findElement(By.linkText(xPath)).click();
								 		driver.findElement(By.xpath(xPath)).click();
									} catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
								 	if(property.equals("BUTTON")){
								 		Thread.sleep(3000);
								 		try {driver.findElement(By.id(xPath)).click();}
								 		catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
								 	
								 	if(property.equals("TEXTFIELD")){
								 		Thread.sleep(3000);
								 		try {driver.findElement(By.id(xPath)).click();}
								 		catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
								 	
								 	if(property.equals("RADIOBUTTON")){
								 		try {driver.findElement(By.xpath(xPath)).click();}
								 		catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
								 
								 	if(property.equals("CHECKBOX")){try {
								 		driver.findElement(By.xpath(xPath)).click();
									} catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
								 	break;
							case "SELECT":
							 		if(property.equals("COMBOBOX")){
							 			try {
							 			Select selectbox = new Select(driver.findElement(By.name(xPath)));
							 			selectbox.selectByVisibleText(value);					 			
							 			//WebElement select= driver.findElement(By.xpath(xPath));
							 			//select.findElement(By.xpath("//option[contains(text(),'" + value + "')]")).click();
							 			/*List<WebElement> options = selectbox.getOptions();
							 			for (WebElement option : options) {
							 				if (option.getText().equals(value)){
							 					option.click();
							 					break;
							 				}}*/
										} catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
							 		break;
							case "DRAGANDDROP":
							 		if(property.equals("DRAGOBJECT")){try {
							 			String[] str_array = xPath.split(",");
							 			String Src = str_array[0];
							 			String Trg = str_array[1];		
							 			WebElement objsource = driver.findElement(By.xpath(Src));
							 			WebElement objtarget = driver.findElement(By.xpath(Trg));
							 			(new Actions(driver)).dragAndDrop(objsource, objtarget).perform();
	
							 		} catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
							 		break;
							 		
							case "CLEAR":
								driver.findElement(By.xpath(xPath)).clear();
								break;

							case "VERIFY":
								 	if(property.equals("PAGETITLE")){
								 		if(driver.getTitle().equals(expectedResult)){
								 			xUpdateTestDetails("Pass");
								 			WriteSheet.addCell(new Label(9,RowNo-1,driver.getTitle()));
								 			WriteSheet.addCell(new Label(10,RowNo-1,"Pass",xFillCell(Colour.GREEN)));}
								 		else {
								 			xUpdateTestDetails("Fail");
								 			WriteSheet.addCell(new Label(9,RowNo-1,driver.getTitle()));
								 			WriteSheet.addCell(new Label(10,RowNo-1,"Fail", xFillCell(Colour.RED)));}}
									if(property.equals("LINK")){if(driver.findElement(By.xpath(xPath)).isDisplayed()){xUpdateTestDetails("Pass");WriteSheet.addCell(new Label(10,RowNo-1,"Pass",xFillCell(Colour.GREEN)));}else {xUpdateTestDetails("Fail");WriteSheet.addCell(new Label(10,RowNo-1,"Fail", xFillCell(Colour.RED)));}}
									if(property.equals("LABLE")){try {
										if(driver.findElement(By.xpath(xPath)).isDisplayed()){
											xUpdateTestDetails("Pass");
											WriteSheet.addCell(new Label(10,RowNo-1,"Pass", xFillCell(Colour.GREEN)));}
										else {
											xUpdateTestDetails("Fail");
											WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));}
									} catch (Exception e) {
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
									}}
									if(property.equals("BUTTON")){if(driver.findElement(By.id(xPath)).isDisplayed()){xUpdateTestDetails("Pass");WriteSheet.addCell(new Label(10,RowNo-1,"Pass",xFillCell(Colour.GREEN)));}else {xUpdateTestDetails("Fail");WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));}}
									if(property.equals("RADIOBUTTON")){if(driver.findElement(By.xpath(xPath)).isDisplayed()){xUpdateTestDetails("Pass");WriteSheet.addCell(new Label(10,RowNo-1,"Pass",xFillCell(Colour.GREEN)));}else {xUpdateTestDetails("Fail");WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));}}
									if(property.equals("CHECKBOX")){if(driver.findElement(By.xpath(xPath)).isDisplayed()){xUpdateTestDetails("Pass");WriteSheet.addCell(new Label(10,RowNo-1,"Pass",xFillCell(Colour.GREEN)));}else {xUpdateTestDetails("Fail");WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));}}
									if(property.equals("TEXTFIELD")){if(driver.findElement(By.xpath(xPath)).isDisplayed()){xUpdateTestDetails("Pass");WriteSheet.addCell(new Label(10,RowNo-1,"Pass",xFillCell(Colour.GREEN)));}else {xUpdateTestDetails("Fail");WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));}}
									break;
							case "SWITCHTO":
									if(property.equals("IFRAME")){try {
									driver.switchTo().frame(driver.findElement(By.id(xPath)));
									//WebElement iframe=driver.findElement(By.tagName("iframe"));
									//driver.switchTo().frame(iframe);
									//System.out.println(iframe);
									}catch (Exception e) { 
									Error= e.getMessage();
									WriteSheet.addCell(new Label(11,RowNo-1,Error));
									xUpdateTestDetails("Fail");
									WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
									xScreenShot();
									break TestCaseLoop;
									}}
									if(property.equals("POPUP")){try {
										Thread.sleep(3000);
										for (String handle : driver.getWindowHandles()) {
											if(!handle.equals(HandleBefore)){
											driver.switchTo().window(handle);
											break;}
											}
										}catch (Exception e) { 
										Error= e.getMessage();
										WriteSheet.addCell(new Label(11,RowNo-1,Error));
										xUpdateTestDetails("Fail");
										WriteSheet.addCell(new Label(10,RowNo-1,"Fail",xFillCell(Colour.RED)));
										xScreenShot();
										break TestCaseLoop;
										}}
									break;
							case "CALL":
									if(Function.equals("Retake")){Initilize.Retake(xPath,value);}
									if(Function.equals("Groups")){Initilize.Groups();}
									break;								
							case "WAIT":
									Thread.sleep(Integer.parseInt(value));
									break;
							case "ENTER":
								driver.findElement(By.xpath(xPath)).sendKeys(Keys.RETURN);
								break;
							case "KEYDOWN":
								driver.findElement(By.xpath(xPath)).sendKeys(Keys.DOWN);
								break;
							
							case "MOUSEOVER":
								Actions builder = new Actions(driver);
								WebElement tagElement = driver.findElement(By.xpath(xPath));
								builder.moveToElement(tagElement).build().perform();
								break;
																
							case "SWITCHBACK":
									if(property.equals("IFRAME")){driver.switchTo().defaultContent();}
									if(property.equals("POPUP")){
										for (String handle : driver.getWindowHandles()) {
											if(handle.equals(HandleBefore)){
											driver.switchTo().window(handle);
											break;}
											}}
									break;
							case "SKIP":
									break;
							case "SCREENSHOT":
									xScreenShot();
									break; 		
							} 
		         		}
		         		
		        	 }
		         }
		   Writebook.write();
		   Writebook.close();
	 }
}
