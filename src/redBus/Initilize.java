package redBus;

import java.io.File;
import org.testng.annotations.Test;
import redBus.PageObjects.UtilityScript;

public class Initilize extends UtilityScript {
	 @Test
	 public void Test_initilize_main() throws Exception {
	  File directory = new File ("C:\\MobileTesting");
	  try{Runtime.getRuntime().exec("wscript.exe "+directory.getCanonicalPath()+"\\initilize.vbs" );
	  }
	    catch(Exception e){e.printStackTrace();
	  }
	  xKillIEs(); 
	  Wait(3000);
	 }
	}
