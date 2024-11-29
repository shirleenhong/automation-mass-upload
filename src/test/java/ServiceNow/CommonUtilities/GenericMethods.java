package ServiceNow.CommonUtilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.ConsoleAppender;
import org.apache.log4j.DailyRollingFileAppender;
import org.apache.log4j.FileAppender;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;
import org.apache.log4j.SimpleLayout;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class GenericMethods {

	protected static Logger log = Logger.getLogger(GenericMethods.class.getSimpleName());

	protected Properties property;
	WebDriver driver;
	
	public static Logger LogFileCreator(String filepath,String classname)
	{
		log= Logger.getLogger(classname);
		DailyRollingFileAppender fileappender = new DailyRollingFileAppender();
	      fileappender.setFile(filepath);
	      fileappender.setDatePattern("'.' yyyy-MM-dd");
	      fileappender.setName("test");
	      
	      PatternLayout filePatternlayout = new PatternLayout();
	      filePatternlayout.setConversionPattern("%d - %c -%p - %m%n");
	      filePatternlayout.activateOptions();
	      
	      fileappender.setLayout(filePatternlayout);
	      fileappender.activateOptions();
	      
	      PatternLayout consolePatternlayout = new PatternLayout();
	      consolePatternlayout.setConversionPattern("%5p [%t] (%F:%L)- %m%n");
	      consolePatternlayout.activateOptions();
	      
	      ConsoleAppender consoleappender = new ConsoleAppender();
	      consoleappender.setLayout(consolePatternlayout);
	      consoleappender.activateOptions();
	     
	      log.addAppender(fileappender);
	      log.addAppender(consoleappender);
	      
	      return log;
	}
	
	public void loadPropertyFile(String filePath){
		File f=new File(filePath);
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e)	 {
			System.out.println(e);
		}
				//new FileInputStream(f);
		property=new Properties();
		try {
			property.load(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	public WebDriver launchBrowserChrome(String URL,String Username,String DriverPath) throws InterruptedException
	{	
		System.setProperty("webdriver.chrome.driver", DriverPath);
	     
		DesiredCapabilities capabilities = DesiredCapabilities.chrome();
		capabilities.setCapability (CapabilityType.ACCEPT_SSL_CERTS, true);
		ChromeOptions options = new ChromeOptions();
		options.addArguments("chrome.switches","--disable-extensions");	
		options.addArguments("disable-infobars");
		options.addArguments("user-data-dir=C:\\Users\\"+Username+"\\AppData\\Local\\Google\\Chrome\\User Data\\test1");
		capabilities.setCapability(ChromeOptions.CAPABILITY, options);
		driver = new ChromeDriver(capabilities); 
		
        driver.get(URL);
        log.info("Site URL Entered successfully...");
        driver.manage().window().maximize();
        return driver;
	}
	
	public static void clickExecute(WebDriver driver, By byObject){
		
		WebDriverWait wait= new WebDriverWait(driver, 900);
		wait.until(ExpectedConditions.elementToBeClickable(byObject));
		/*(new WebDriverWait(driver, 20))
				  .until(ExpectedConditions.elementToBeClickable(byObject));*/
		((JavascriptExecutor)driver).executeScript("arguments [0].click();", driver.findElement(byObject));
	}

	public void waitForPageLoaded() throws Exception {
//		ExpectedCondition<Boolean> expectation = new
//				ExpectedCondition<Boolean>() {
//			public Boolean apply(WebDriver driver) {
//				return ((JavascriptExecutor) driver).executeScript("return document.readyState").toString().equals("complete");
//			}
//		};
//		try {
//			Thread.sleep(4000);
//			WebDriverWait wait = new WebDriverWait(driver, 500);
//			wait.until(expectation);
//		} catch (Exception error) {
//			Assert.fail("Timeout waiting for Page Load Request to complete.");
//			//failureScreenShot(driver);
//		}
		
		  JavascriptExecutor js = (JavascriptExecutor)driver;


		  //Initially bellow given if condition will check ready state of page.
		  if (js.executeScript("return document.readyState").toString().equals("complete")){ 
		   return; 
		  } 

		  //This loop will rotate for 60 times to check If page Is ready after every 1 second.
		  //You can replace your value with 60 If you wants to Increase or decrease wait time.
		  for (int i=0; i<60; i++){ 
		   try {
		    Thread.sleep(1000);
		    }catch (InterruptedException e) {} 
		   //To check page ready state.
		   if (js.executeScript("return document.readyState").toString().equals("complete")){ 
		    break;
		   }   
		  }
		 
	}
	
	public void logInfo(String logData){
		BasicConfigurator.configure();
		FileAppender fileAppender= new FileAppender();
		fileAppender.setFile("C:\\PCS RPA\\LogFile\\SNInvoicingLogs.txt");
		fileAppender.setLayout(new SimpleLayout());
		log.addAppender(fileAppender);
		fileAppender.activateOptions();
		log.info(logData);
	}

}
