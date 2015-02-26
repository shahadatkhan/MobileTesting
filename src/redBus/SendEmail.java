/*###################################################################################################
Name: SendEmail.java 
Description: Calls senemail.vbs file to send the test execution results email
Author: Sudheer Bandi
Date Created:01/17/2014
Date Modified:
Updated Comments:
#####################################################################################################*/

package redBus;

import java.io.File;

import org.testng.annotations.Test;

import redBus.PageObjects.UtilityScript;

public class SendEmail extends UtilityScript {
	@Test
	public void Test_SendEmail_main() throws InterruptedException {
		Wait(3000);
		File directory = new File(".//");
		// Print(MethodName);
		try {
			Runtime.getRuntime().exec("wscript.exe " + directory.getCanonicalPath()+ "\\sendemail.vbs " + MethodName);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}