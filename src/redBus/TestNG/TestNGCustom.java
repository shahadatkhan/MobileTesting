package redBus.TestNG;

import org.testng.*;
import org.testng.TestListenerAdapter;

public class TestNGCustom extends TestListenerAdapter {

 // Take screen shot only for failed test case
 @Override
 public void onTestFailure(ITestResult tr){
  try {
	  redBus.PageObjects.UtilityScript.xScreenShot();
  } catch (Exception e1) {
   // TODO Auto-generated catch block
   e1.printStackTrace();
  }
  try {
	  redBus.PageObjects.UtilityScript.xUpdateTestDetails("FAIL");
  } catch (Exception e) {
   // TODO Auto-generated catch block
   e.printStackTrace();
  }
 }
 @Override
 public void onTestSkipped(ITestResult tr) {
  // p2pZions.PageObjects.UtilityScript.xScreenShot();
  try {
	  redBus.PageObjects.UtilityScript.xUpdateTestDetails("SKIPPED");
  } catch (Exception e) {
   // TODO Auto-generated catch block
   e.printStackTrace();
  }  
 }

 @Override
 public void onTestSuccess(ITestResult tr) {
  //p2pZions.PageObjects.UtilityScript.xScreenShot();
  try {
	  redBus.PageObjects.UtilityScript.xUpdateTestDetails("PASS");
  } catch (Exception e) {
   // TODO Auto-generated catch block
   e.printStackTrace();
  }
 }
 }