package redBus.PageObjects;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.net.InetAddress;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import javax.imageio.ImageIO;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.Pattern;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import org.apache.commons.io.FileUtils;
import org.testng.Reporter;

public class UtilityScript extends TestData {

 // Get date time
 public java.lang.String xGetDateTime() throws Exception {
  // get current date time with Date() to create unique file name
  DateFormat dateFormat = new SimpleDateFormat("hh_mm_ssaadd_MMM_yyyy");
  // get current date time with Date()
  Date date = new Date();
  return (dateFormat.format(date));
 }

 // DateFormat = "MMM dd, yyyy";
 public java.lang.String xGetDate(String DateFormat) throws Exception {
  // get current date time with Date() to create unique file name
  DateFormat dateFormat = new SimpleDateFormat(DateFormat);
  // get current date time with Date()
  Date date = new Date();
  return (dateFormat.format(date));
 }

 // Get date time with SelText
 public java.lang.String xGetDateTimeSel() throws Exception {
  // get current date time with Date() to create unique file name
  DateFormat dateFormat = new SimpleDateFormat("hh_mm_ssaadd_MMM_yyyy");
  // get current date time with Date()
  Date date = new Date();
  return ("S_" + dateFormat.format(date));
 }

 // Get date time with System IP
 public java.lang.String xGetDateTimeIP() throws Exception {
  // get current date time with Date() to create unique file name
  DateFormat dateFormat = new SimpleDateFormat("hh_mm_ssaa_dd_MMM_yyyy");
  // get current date time with Date()
  Date date = new Date();
  // To identify the system
  InetAddress ownIP = InetAddress.getLocalHost();
  return (dateFormat.format(date) + "_IP" + ownIP.getHostAddress());
 }

//Get browser to run the test
 public static String xBrowser() throws Exception {
	 Workbook workbook = Workbook.getWorkbook(new File("C:\\MobileTesting\\TestSetup\\TestSetup.xls"));
	 Sheet sheet = workbook.getSheet("TestSetup");
	 String browser = sheet.getCell(1,6).getContents();
	  return (browser);
	 }
//Get Testpath to run the test
public static String xTestPath() throws Exception {
	 Workbook workbook = Workbook.getWorkbook(new File("C:\\MobileTesting\\TestSetup\\TestSetup.xls"));
	 Sheet sheet = workbook.getSheet("TestSetup");
	 String TestPath = sheet.getCell(1,7).getContents();
	  return (TestPath);
	 }

//Get TestResult to run the test
public static String xTestResult() throws Exception {
	 Workbook workbook = Workbook.getWorkbook(new File("C:\\MobileTesting\\TestSetup\\TestSetup.xls"));
	 Sheet sheet = workbook.getSheet("TestSetup");
	 String TestResult = sheet.getCell(1,8).getContents();
	  return (TestResult);
	 }

//Get TestSuite name
public static String xTestSuite() throws Exception {
	 Workbook workbook = Workbook.getWorkbook(new File("C:\\MobileTesting\\TestSetup\\TestSetup.xls"));
	 Sheet sheet = workbook.getSheet("TestSetup");
	 String TestSuite = sheet.getCell(1,4).getContents();
	  return (TestSuite);
	 }

 
 public static void Wait(int MilliSec) throws InterruptedException {
  Thread.sleep(MilliSec);
 }

 public void Print(String Text) {
  System.out.println(Text);
  Reporter.log(Text);
  String Temp = Text;
  sMessages = sMessages + Temp.replaceAll(" ", "_") + "#";
  //System.out.println(Temp);
  //System.out.println(sMessages);
 }

 public java.lang.String xAddMinutesToTheDateTime(String Date_TimeFormat,
   int NumberOfMinutes) throws InterruptedException, ParseException {
  SimpleDateFormat sdf = new SimpleDateFormat(DateTimeFormat);
  Calendar c = Calendar.getInstance();
  c.setTime(sdf.parse(Date_TimeFormat));
  c.add(Calendar.MINUTE, NumberOfMinutes); // number of minutes
  String str = sdf.format(c.getTime());
  String delimiter = "_";
  String[] temp;
  temp = str.split(delimiter);
  for (int i = 0; i < temp.length - 1; i++) {
   NewDate = temp[i];
   NewTime = temp[i + 1];
  }
  // Print(NewDate);
  // Print(NewTime);

  return (str); // dt is now the new date
 }

 public java.lang.String xAddDaysToTheDateTime(String CurrentDate,
   int NumberOfDays, String DateFormat)
   throws InterruptedException, ParseException {
  SimpleDateFormat sdf = new SimpleDateFormat(DateFormat);
  Calendar c = Calendar.getInstance();
  c.setTime(sdf.parse(CurrentDate));
  c.add(Calendar.DATE, NumberOfDays); // number of days
  String str = sdf.format(c.getTime());
  return (str); // dt is now the new date
 }

 public java.lang.String xGetCurrentDateEST(String DateFormat) throws Exception {
  SimpleDateFormat dateFormat = new SimpleDateFormat(DateFormat);
  dateFormat.setTimeZone(TimeZone.getTimeZone("EST5EDT"));
  NewDate = dateFormat.format(new Date());
  return (NewDate);
 }

 public java.lang.String xGetCurrentTimeEST() throws Exception {
  SimpleDateFormat dateFormat = new SimpleDateFormat("hh:mm a");
  dateFormat.setTimeZone(TimeZone.getTimeZone("EST5EDT"));
  NewTime = dateFormat.format(new Date());
  return (NewTime);
 }

 public void xKillIEs() throws Exception {
  Wait(3000);
   File directory = new File("C:\\MobileTesting\\");
  try {
   Runtime.getRuntime().exec("wscript.exe " + directory.getCanonicalPath() + "\\KillIEs.vbs");
  } catch (Exception e) {
   e.printStackTrace();
  }
  Wait(5000); 
 }
 
 public void xKillExcel() throws Exception {
	  Wait(3000);
	   File directory = new File("C:\\MobileTesting\\");
	  try {
	   Runtime.getRuntime().exec("wscript.exe " + directory.getCanonicalPath() + "\\KillExcel.vbs");
	  } catch (Exception e) {
	   e.printStackTrace();
	  }
	  Wait(5000); 
	 }

 public boolean xFileExist(String FileNameWithPath) throws Exception {
  java.io.File myDir = new java.io.File(FileNameWithPath);
  if (myDir.exists()) {
   Print("file exist");
   return true;
  } else {
   Print("file does not exist");
   assertTrue(false);
   return false;
  }
 }

 public void xMakeFileCopy(String NewFileNameWithPath,
   String FileNameWithPath) throws Exception {
  java.io.File base = new java.io.File(FileNameWithPath);
  java.io.File newfile = new java.io.File(NewFileNameWithPath);
  if (xFileExist(FileNameWithPath)) {
   FileUtils.copyFile(base, newfile);
  } else {
   Print("file does not existcould not copy");
   assertTrue(false);
  }
  if (xFileExist(NewFileNameWithPath)) {
   Print("file copied sucessfully");
  }

 }

 
 /*
 public void xDeleteFile(String FileNameWithPath) throws Exception {
  java.io.File file = new java.io.File(FileNameWithPath);
  if (xFileExist(FileNameWithPath)) {
   FileUtils.deleteQuietly(file);
   Print("File Deleted Successfully");
  } else {
   Print("file does not exist.Could not Delete");
   // assertTrue(false);
  }
 }
*/
 
 public static void xScreenShot() {
  try {

   String NewFileNamePath;
   java.awt.Dimension resolution = Toolkit.getDefaultToolkit()
   .getScreenSize();
   Rectangle rectangle = new Rectangle(resolution);

   // Get the dir path
   File directory = new File("C:\\MobileTesting\\");
   // System.out.println(directory.getCanonicalPath());

   // get current date time with Date() to create unique file name
   DateFormat dateFormat = new SimpleDateFormat(
     "MMM_dd_yyyy__hh_mm_ssaa");
   // get current date time with Date()
   Date date = new Date();
   // System.out.println(dateFormat.format(date));

   // To identify the system
   InetAddress ownIP = InetAddress.getLocalHost();
   // System.out.println("IP of my system is := "+ownIP.getHostAddress());

   NewFileNamePath = directory.getCanonicalPath() + "\\ScreenShots\\" + Method + "___" + dateFormat.format(date) + "_"  + ownIP.getHostAddress() + ".png";
   System.out.println(NewFileNamePath);

   // Capture the screen shot of the area of the screen defined by the
   // rectangle
   Robot robot = new Robot();
   //BufferedImage bi = new BufferedImage();
   BufferedImage bi = robot.createScreenCapture(new Rectangle(rectangle));
   ImageIO.write(bi, "png", new File(NewFileNamePath));
   NewFileNamePath = "<a href=" + NewFileNamePath + ">ScreenShot"
     + "</a>";
   // Place the reference in TestNG web report
   Reporter.log(NewFileNamePath);

  } catch (AWTException e) {
   e.printStackTrace();
  } catch (IOException e) {
   e.printStackTrace();
  }
 }
 
 public static void xUpdateTestDetails(String Status) throws Exception   {
  File directory = new File("C:\\MobileTesting\\");
  String Temp = Method + "_" + Status;
  if (Method != ""){
   try {
    Runtime.getRuntime().exec( "wscript.exe " + directory.getCanonicalPath() + "\\UpdateTestDetails.vbs "+ Temp + " " + sMessages);
    Method = "";
    sMessages = "";
   } catch (Exception e) {
    e.printStackTrace();
   }
  }
  Wait(5000); // Allow OS to kill the process
 }  
 
 public WritableCellFormat xFillCell(Colour colour) throws WriteException{
	 	    WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
		    WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
		    cellFormat.setBackground(colour);
		    return cellFormat;
	}
 
 public WritableCellFormat xFormatCell(Pattern pattern) throws WriteException{
	    WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 10);
	    WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	    cellFormat.setBackground(Colour.BLACK,pattern);
	    return cellFormat;
}
}