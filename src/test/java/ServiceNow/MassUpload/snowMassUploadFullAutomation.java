package ServiceNow.MassUpload;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.TestListenerAdapter;
import org.testng.TestNG;
import org.testng.annotations.Test;

import ServiceNow.CommonUtilities.ExcelFunctions;
import ServiceNow.CommonUtilities.GenericMethods;

@SuppressWarnings("deprecation")
public class snowMassUploadFullAutomation extends GenericMethods 
{
	WebDriver driver;
	String SNInvoicingURL;
	String Username;
	String driverPath;
	public static XSSFWorkbook workbook;
	int matchfound=0 ,iExcelRow=0;
	String Enviornment;
	String ConfigurationFilePath; 
	int intLastRowForProcessing = 1;
	protected static Logger SNOWLog;
	String IsUploadAttachmentRequired="No"; 
	String InputDataFile;
	String InputSheetname;
	String CurrentCountry="";
	String currentGroup="";
	String NextCountry = "",NextRowStatus = "",Nextgroup = "";
	boolean IsaddingInvoiceflag=false, restartFromFailedRow = false;
	boolean IsProcessingLastRecordRow,IsDropdownfieldvalueFound, startRun;
	WebDriverWait wait;
	Boolean flag = false;
	String ReportRows[];
	int InvoiceNumber = 2;
	String IsApplicationRunInMinimisedMode ;
	String RequestTemplateName,GU,Country,BillingRequestType,FormType,ContractNumber,InvoiceCurrency,DisplayTimeOfSupply,IsTimeOfSupplyDifferent,TimeOfSupply,PO,InvoiceDate,ServiceRenderedDate,DiaplayQuantity;
	String WBSE,LineItem,MaterialSalesText,AdditionalInformation,Quantity,Amount,ContractualObligation,AmountOfExpence,PaymentTerms,PaymentMethod,FormHeader,Language;
	String AddValueDays,FixedValueDate,Logo,Layout,ShouldBillingTeamDeliverInvoice,DeliveryMethod,CoverLetterNeeded,IncludeAttachment,AttachmentRequired;
	String IdentificationNumber,AdditionalNote1,AdditionalNote2,CustomerAddInfo,CompanyAddInfo,ReviewerId,ApproverId,AttachmentPath,WebPortal,Email,Group,RPAStatus;
	String SupplementalDocument,PleaseProvideExplanation,Comments,BillToParty,PostalAddress;
	int print, invoiceCntr, negTest, test, testFlag, countryFlag, CountryColumn = 2,GroupColumn = 50,RPAStatusColumn = 51,TimeTakenToCreateInvoice = 52,RequestNumbercolumn = 53, RITMNumbercolumn = 54;
	double startTime,EndTime,TotalTimeTaken;

	
	public void ExcelCurrentRowParameterSetting(String ReportRows[])
	{
		RequestTemplateName = ReportRows[0];
		GU = ReportRows[1];
		Country = ReportRows[2];
		BillingRequestType = ReportRows[3];
		FormType = ReportRows[4];
		ContractNumber = ReportRows[5];
		InvoiceCurrency = ReportRows[6];
		DisplayTimeOfSupply = ReportRows[7];
		IsTimeOfSupplyDifferent = ReportRows[8];
		TimeOfSupply = ReportRows[9];
		PO = ReportRows[10];
		InvoiceDate =  ReportRows[11];
		ServiceRenderedDate = ReportRows[12];
		DiaplayQuantity = ReportRows[13];
		Language = ReportRows[14];
	    WBSE = ReportRows[15]; 
	    LineItem = ReportRows[16];
	    MaterialSalesText = ReportRows[17];
	    AdditionalInformation = ReportRows[18];
	    Quantity =  ReportRows[19];
	    Amount = ReportRows[20];
	    ContractualObligation = ReportRows[21];
	    AmountOfExpence = ReportRows[22];
	    PaymentTerms = ReportRows[23];	    
		AddValueDays = ReportRows[24];
		FixedValueDate =  ReportRows[25];
		PaymentMethod = ReportRows[26];
		FormHeader = ReportRows[27];
		AdditionalNote1 = ReportRows[28];
		AdditionalNote2 = ReportRows[29];
		CustomerAddInfo = ReportRows[30]; 
		CompanyAddInfo = ReportRows[31];
		Logo = ReportRows[32];
		BillToParty = ReportRows[33];
		Layout = ReportRows[34];
		ShouldBillingTeamDeliverInvoice = ReportRows[35];
		DeliveryMethod =  ReportRows[36];
		WebPortal = ReportRows[37];
		Email = ReportRows[38];
		PostalAddress = ReportRows[39];
		CoverLetterNeeded =  ReportRows[40];
		IncludeAttachment = ReportRows[41];
		AttachmentRequired = ReportRows[42];
		SupplementalDocument = ReportRows[43];
		PleaseProvideExplanation = ReportRows[44];
		IdentificationNumber = ReportRows[45];
		ReviewerId = ReportRows[46];
		ApproverId =  ReportRows[47];
		AttachmentPath = ReportRows[48];
		Comments = ReportRows[49];	
		Group = ReportRows[50];
		RPAStatus = ReportRows[51];			
		
	}
	
	@Test(priority=1)
	public void ReadConfiguration() throws Exception
	{
			try{
			//Userinput file Configuration
			loadPropertyFile("C:\\PCS RPA\\DigitalInvoicing\\Config\\Configuration Files\\UserInputFullAutomation.properties");
			SNOWLog.info("Userinput-fullAutomation property file configured successfully.");
			}catch(Exception e)
			{
				SNOWLog.info("Error in Reading UserInputFullAutomation property file. Check availability of the file in 'C:\\PCS RPA\\DigitalInvoicing\\Config\\Configuration Files\\' folder and restart the automation.");
				SNOWLog.info(e.getLocalizedMessage());
				Runtime.getRuntime().exec(new String[] {"wscript.exe","C:\\PCS RPA\\DigitalInvoicing\\ServerFiles\\" + "SendGenerateInvoicesErrorMail.vbs"});
				throw e;
			}
			//Configuration file set up
			ConfigurationFilePath=property.getProperty("ConfigFileName");
			SNOWLog.info("SNOW Base Configuration excel file path : "+ConfigurationFilePath);
			File ConfigFile = new File(ConfigurationFilePath);
			if(!ConfigFile.exists())
			{
				SNOWLog.info("SNOW mass Invoice-ticket upload Base Configuration excel file does not exist at the specified location.Check for the availabilty of file and try running the BOT.");
			//	tearDown();
			}
			//Username set up
			Username=property.getProperty("username");
			SNOWLog.info("Username from the userinput-full automation property file : "+Username );
			try
			{
				//Config-Excel Path setting and Reading
				ExcelFunctions.setPath(ConfigurationFilePath, "Configuration");
				SNOWLog.info("Config-Excel Path set successfully.");
				
				IsUploadAttachmentRequired = ExcelFunctions.getData(12, 6);
				SNOWLog.info("Is Upload Attachment Required: "+IsUploadAttachmentRequired);
				
				Enviornment = ExcelFunctions.getData(6, 6);
				SNOWLog.info("Service-now tool running enviornment: "+Enviornment);
				
				if(Enviornment.equalsIgnoreCase("Test"))
				  SNInvoicingURL = ExcelFunctions.getData(7, 6);
				else
				   SNInvoicingURL = ExcelFunctions.getData(8, 6);	
				
				IsApplicationRunInMinimisedMode = ExcelFunctions.getData(13, 6);
				SNOWLog.info("Is application should run in minimised mode : "+IsApplicationRunInMinimisedMode);
		}
		catch(Exception e)
		{
			SNOWLog.info("Error in ReadConfiguration module.Please restart automation.");
			SNOWLog.info(e.getLocalizedMessage());
			Runtime.getRuntime().exec(new String[] {"wscript.exe","C:\\PCS RPA\\DigitalInvoicing\\ServerFiles\\" + "SendGenerateInvoicesErrorMail.vbs"});
			//tearDown();
			throw e;
		}
	}
	
	@Test(priority=2, dependsOnMethods = { "ReadConfiguration" })
	public void setUp() throws Exception{
		try{
		SNOWLog.info("Service-now application URL : "+SNInvoicingURL);
		String chromeDriver = property.getProperty("chromeDriver");
		SNOWLog.info("ChromeDriver executable file path: "+chromeDriver);		
		File chromeDriverfile = new File(chromeDriver);
		if(!chromeDriverfile.exists())
		{
			SNOWLog.info("chromedriver.exe file does not exist at the specified location.Check for the availabilty of file and try running the BOT.");
			throw new Exception("chromedriver.exe file does not exist at the specified location.Check for the availabilty of file and try running the BOT.");
		}
		driver = launchBrowserChrome(SNInvoicingURL,Username,chromeDriver);
		SNOWLog.info("Navigated to the site successfully!");
		wait = new WebDriverWait(driver, 60); //Dynamic wait declaration
		waitForPageLoaded();
		}catch(Exception e)
		{
			SNOWLog.info("Unable to launch browser.Please restart the bot.");
			SNOWLog.info(e.getLocalizedMessage());
			throw new Exception("Unable to launch browser.Please restart the bot.");
		}
	}
	
	@Test (priority=3, dependsOnMethods = { "setUp" })
	public void InvoicingDataInsertion() throws Exception{
		 SNOWLog.info("Into the InvoicingDataInsertion module.");
		 waitForPageLoaded();
		 
		 String WantToSubmitReport = ExcelFunctions.getData(11, 6);
		 InputDataFile = ExcelFunctions.getData(9, 6);
		 SNOWLog.info("SNOW Invoice data excel sheet file path : "+InputDataFile);
		 File InputDataExcelpath = new File(InputDataFile);
		 if(!InputDataExcelpath.exists())
		 {
			 SNOWLog.info("'InvoiceData_FullAutomation_Pilot.xlsm' file does not exist at the specified location.Check for the availabilty of file and try running the BOT.");
			 tearDown();
		 }
		 InputSheetname = ExcelFunctions.getData(10, 6);
		 SNOWLog.info("SNInvoicing input data sheet source name: "+InputSheetname);
			//Excel sheet reading
			ExcelFunctions.setPath(InputDataFile,InputSheetname);
			
			//Restarting logic
			String LastRecordToStartForecasting = ExcelFunctions.getData(0,56);
			if(!LastRecordToStartForecasting.equalsIgnoreCase(""))
				intLastRowForProcessing=Integer.parseInt(LastRecordToStartForecasting);
			SNOWLog.info("The excel-row considered for successful ticket generation: "+intLastRowForProcessing);
			
			String InvoicingData[] = ExcelFunctions.readData(InputDataFile, InputSheetname);
			SNOWLog.info("Total number of records for invoice generation in the excel sheet are: "+InvoicingData.length);					
			
			for(test=1; test<InvoicingData.length;test++ )
			{
				String CurrentRow[] = InvoicingData[test].split("--");
				String PrevRow[] = InvoicingData[test-1].split("--");
				String PrevCountryCheck = PrevRow[CountryColumn];
				String CurrentCountryCheck = CurrentRow[CountryColumn];
								
				if (CurrentCountryCheck.equalsIgnoreCase(PrevCountryCheck)) {
					testFlag++;
				}
			}
			
			Boolean startRun = true;
			
		 try 
			{		
			ArrayList<Integer> cntFlag = new ArrayList<Integer>();
			
	        //For loop for the contract should start 
			for( iExcelRow=intLastRowForProcessing; iExcelRow<=InvoicingData.length;iExcelRow++)
			{
				int iGridRowIterator=1;
				
				if(iExcelRow==InvoicingData.length)
					flag=true;
				
				if (startRun=true) 
				{
			        startTime = System.currentTimeMillis();
			        SNOWLog.info("The start time of the invoice is: "+startTime);
				}
				
				//Checking for the empty row
				if(!InvoicingData[iExcelRow].isEmpty())
				{
					SNOWLog.info("Current excel row number considred for processing from the excel sheet: " + iExcelRow );
					String	ReportRows[] = InvoicingData[iExcelRow].split("--");
					ExcelCurrentRowParameterSetting(ReportRows); //Setting all report Parameters
					
					currentGroup  = Group;
					SNOWLog.info("The Group for the current excel row: "+currentGroup);
					CurrentCountry = Country;
					SNOWLog.info("The country for the current excel row: "+CurrentCountry);
					NextRowStatus="";
					
					/**
					 * TO DO
					 */
					
//			        startTime = System.currentTimeMillis();
//			        SNOWLog.info("The start time of the invoice is: "+startTime);
			        
					if ((DeliveryMethod.equalsIgnoreCase("IDOC")&&(!((Country.equalsIgnoreCase("Norway"))||(Country.equalsIgnoreCase("Sweden"))||(Country.equalsIgnoreCase("Denmark"))||(Country.equalsIgnoreCase("Finland"))))))
						{
							ExcelStatusWriter("Delivery Method: "+DeliveryMethod+" is not available for "+CurrentCountry, true);
						}
					
					else
					{
					//checking for the status 
					if((RPAStatus.contains("Failure")||RPAStatus.equalsIgnoreCase("Create Invoice")))
					{
						if((!currentGroup.equalsIgnoreCase(Nextgroup))||(restartFromFailedRow==true))
						{
							SNOWLog.info("Excel current row group: "+currentGroup + " Excel next row group: " + Nextgroup);
							//Logic for create invoice request template
							if(IsaddingInvoiceflag==false)
							{
								 InvoiceNumber = 2;	
						        clickExecute(driver, By.id(property.getProperty("TransactionName")));         //Transaction name
						        
						        //Minimised version implementation
								if( IsApplicationRunInMinimisedMode.equalsIgnoreCase("yes"))
								{
									driver.manage().window().setPosition(new Point(-2000, 0));
									SNOWLog.info("Running SNOW mass Invoice ticket upload application in Minimized view" );	
								}
								
								//TODO: Input Request Template Name, Geo Unit, and Country
						        String TemplateName = RequestTemplateName + "_RPA";
						        InsertTextFieldData("Transaction name",TemplateName,"TransactionName");
						       
//						        startTime = System.currentTimeMillis();
//						        SNOWLog.info("The start time of the invoice is: "+startTime);
//						        waitForPageLoaded();
						        
						        negTest = 1;
						        
						        try {
						        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("geographicUnits")));
						        }
						        catch(Exception e)
						        {
						        	Thread.sleep(5000);
						        }
						        
								//Geographic unit
						        try{
								new Select(driver.findElement(By.id(property.getProperty("GeographicUnit")))).selectByVisibleText(GU);
								SNOWLog.info("Geographical unit entered sucessfully.Entered value from input excel sheet is: "+GU);
								waitForPageLoaded();
								Thread.sleep(5000); //Necesasry for country dd to get populate
						        }catch(Exception e)
						        {
						        	SNOWLog.info("Geographic unit dropdown value is not found.Will skip this and move ahead.");
									ExcelStatusWriter("Failure : Geographic unit dropdown field value is not found.",true);
									restartFromFailedRow = true;
									SkippingRowBasedOnCountryAndGroup(InvoicingData);
									continue;
						        }
						        
						        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("country")));
						        
								//Country
								try{
								new Select(driver.findElement(By.id(property.getProperty("Country")))).selectByVisibleText(Country);
								SNOWLog.info("Country entered sucessfully.Entered value from input excel sheet is: "+Country);
								}catch(Exception e)
								{
									SNOWLog.info("Country dropdown svalue is not found.Will skip this and move ahead.");
									ExcelStatusWriter("Failure : Country dropdown field value is not found.",true);
									restartFromFailedRow = true;
									SkippingRowBasedOnCountryAndGroup(InvoicingData);
									continue;
								}
								
						        //Request template button click
								clickExecute(driver, By.xpath(property.getProperty("RequestTemplateButton")));
						        SNOWLog.info("Create Request template Button clicked successfully.");
	        
							}						
					
				    		if((GU.equalsIgnoreCase("ASG"))||(GU.equalsIgnoreCase("Nordics")))
					        {
				    			wait.until(ExpectedConditions.alertIsPresent());
//					        	Thread.sleep(10000);
					        	PopupHandle();
					        }
							
							//TODO: Start of Invoice Input Page
					        iGridRowIterator=1;
//					        Thread.sleep(7000);
					        
					        //cntFlag to capture RITM of requests with same Country
					        //
				    		cntFlag.add(iExcelRow);
				   					        	    		
				    		try
							{
					        	driver.switchTo().defaultContent();
						        Thread.sleep(3000);
						        WebDriverWait wait = new WebDriverWait(driver,60);
						        wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.tagName("iframe")));	
								driver.findElement(By.xpath("//button[contains(@onclick,'termsAccept')]")).click();
								SNOWLog.info("Terms and Conditions accepted.");
							}

							catch(Exception e)
							{
								SNOWLog.info("Input Primary Invoice Information");
							}
				    		
							//Moving to the next page and Logic for Entering primary invoice information			        	
					        SNOWLog.info("Trying to find element : Billing request type on to the webpage."); 
					        //wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.id("gsft_main"))); 
					        
					        wait.until(ExpectedConditions.elementToBeClickable(By.id("IO:1964e6eddb008b00f945f9a41d96196a")));
					        new Select(driver.findElement(By.id("IO:1964e6eddb008b00f945f9a41d96196a"))).selectByVisibleText(BillingRequestType);
					        
					        //Invoice Type
					        IsDropdownfieldvalueFound= InsertMandatorydropDownFieldData("Invoice type",BillingRequestType,"InvoiceType",InvoicingData);
					        if(IsDropdownfieldvalueFound == false)
					        	continue;
					        
//					        if(Country.equalsIgnoreCase("France")||Country.equalsIgnoreCase("Mauritius"))
//					        	PopupHandle();

					        if(Country.equalsIgnoreCase("France")||Country.equalsIgnoreCase("Mauritius"))
					        {
				    			wait.until(ExpectedConditions.alertIsPresent());
//					        	Thread.sleep(10000);
					        	PopupHandle();
					        }
					        
					        //Form type
					        String FormTypeapplication = driver.findElement(By.id(property.getProperty("FormType"))).getText();
					        IsDropdownfieldvalueFound= InsertMandatorydropDownFieldData("Form Type",FormType,"FormType",InvoicingData);
					        if(IsDropdownfieldvalueFound == false)
					        	continue;
					        if(!FormTypeapplication.equalsIgnoreCase(FormType))
					        	PopupHandle();
					         					 
					        //Contract no
					        try{
					        BigDecimal val=new BigDecimal(ContractNumber); 
					        InsertTextFieldData("Contract number",val.toPlainString(),"ContractNo");
					        }catch(Exception e)
					        {
					        	SNOWLog.info("Contract number is in plain format.");
					        	InsertTextFieldData("Contract number",ContractNumber,"ContractNo");
					        }
			
					        //Waiting for the contract details to get populated container click
					        driver.findElement(By.id("IO:52a7e6e1db408b00f945f9a41d961949")).click();
					        PopupHandle();
					        PopupHandle(); //Added for the nordics region

					        String ContractName = driver.findElement(By.id("IO:52a7e6e1db408b00f945f9a41d961949")).getAttribute("value");
			     			SNOWLog.info("Contract name extracted from the appication is: "+ContractName);
					        
			     			int counter = 0;
			     			while(ContractName.isEmpty()&&counter<=30)
					        {
					        	Thread.sleep(1000);
					        	SNOWLog.info("Waiting for the contract details to get populated...");
					        	ContractName = driver.findElement(By.id("IO:52a7e6e1db408b00f945f9a41d961949")).getAttribute("value");
					        	counter++;
					        }
			     			if(ContractName.isEmpty())
			     			{
			     				SNOWLog.info("Bot is unable to retrieve the contract details,Moving to the next record.");
			     				ExcelStatusWriter("Failure : Contract Details not found",true);
			     				restartFromFailedRow = true;
						    	BackButtonClick();
			     				continue;
			     			}
			     						        
					        //Invoice Currency
			     			IsDropdownfieldvalueFound= InsertMandatorydropDownFieldData("Currency",InvoiceCurrency,"InvoiceCurrency",InvoicingData);
					        if(IsDropdownfieldvalueFound == false)
					        	continue;
					        PopupHandle();
			     			
					      //If Contract number checkbox is present then click
					        try{
					        	driver.findElement(By.id(property.getProperty("ContractCheckBox"))).click();
					        	SNOWLog.info("The check box after the contract is clicked successfully");
					        } catch(Exception e)
					        {
					        	SNOWLog.info("The check box after the contract is not present.");
					        }
					        
					        //Clicking on PO 
					        if(PO.equalsIgnoreCase("BLANK"))
					        	EraseFieldData("PO");
					        else if(!PO.equalsIgnoreCase("NA")&&(PO.contains(".")))
					        {
					        	BigDecimal POval=new BigDecimal(PO);
					        	clickExecute(driver, By.id(property.getProperty("PO"))); 
						        InsertTextFieldData("PO number",POval.toPlainString(),"PO");
						    }
					        else if(!PO.equalsIgnoreCase("NA"))
					        {
					        	SNOWLog.info("As PO number is only numeric.");
					        	clickExecute(driver, By.id(property.getProperty("PO")));
						        InsertTextFieldData("PO number",PO,"PO");
					        }
					        
					        //Display-Time of supply 
					        InsertOptionaldropdownFieldData("Display-Time of supply",DisplayTimeOfSupply,"DisplayTimeOfSupply");
					        PopupHandle();
					      
						    //Is the Time of Supply different to the Invoice date?
					        InsertOptionaldropdownFieldData("Is Time Of Supply Different to the invoice date",IsTimeOfSupplyDifferent,"IsTimeOfSupplyDifferent");
						        
						     //Time of supply
					        InsertTextFieldData("Time of supply",TimeOfSupply,"TimeOfSupply");
						       
					        //Date
						    if(InvoiceDate.equalsIgnoreCase("BLANK")) 
						    	EraseFieldData("Date");
						    else if(!InvoiceDate.equalsIgnoreCase("NA"))
						    {
						    	InsertTextFieldData("Date",InvoiceDate,"Date");
					        	waitForPageLoaded();
						        driver.findElement(By.id(property.getProperty("DisplayQuantity"))).click();
						        PopupHandle();
						    }
						    
					      //Service rendered Date - 535099 bug fixed
						    if(ServiceRenderedDate.equalsIgnoreCase("BLANK")) 
						    	EraseFieldData("ServiceRenderedDate");
						    else if(!ServiceRenderedDate.equalsIgnoreCase("NA"))
						    	InsertTextFieldData("Service Rendered Date",ServiceRenderedDate,"ServiceRenderedDate");
					        	
					        //Display Quantity
					        if(!DiaplayQuantity.equalsIgnoreCase("NA"))
						    {
					        	InsertOptionaldropdownFieldData("Display quantity",DiaplayQuantity,"DisplayQuantity");
						        waitForPageLoaded();
				     			SNOWLog.info("Entered Primary Information " );
						    }
					        
			     			//Language
							driver.findElement(By.id(property.getProperty("Language"))).sendKeys(Language);
							SNOWLog.info("Language entered successfully.Entered value from input excel sheet is: "+Language);
						}
						iGridRowIterator++;
						String chariGridRowIteratorval = Integer.toString(iGridRowIterator);
						String GridRowXpathValue,NewXpathValue;
						SNOWLog.info("The Grid Row considered for the updation from dynamic grid:"+iGridRowIterator);
						//Logic for Entering grid row for current excel line item
						Thread.sleep(7000); //Necessary for the WBS information to get populate 
						//WBS
						try{
						GridRowXpathValue = property.getProperty("WBS");
						NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
						wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(NewXpathValue)));
						new Select (driver.findElement(By.xpath(NewXpathValue))).selectByVisibleText(WBSE);
						SNOWLog.info("WBS entered.Entered value from input excel sheet is: " + WBSE);
						}catch(Exception ee)
						{
							try{
								GridRowXpathValue = property.getProperty("WBS");
								NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
								wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(NewXpathValue)));
								new Select (driver.findElement(By.xpath(NewXpathValue))).selectByVisibleText(WBSE);
								SNOWLog.info("WBS entered.Entered value from input excel sheet is: " + WBSE);
								}catch(Exception e)
								{
									SNOWLog.info("WBS Not Found...Please enter the correct WBS.. ");
									ExcelStatusWriter("Failure : WBS Not Found",false);
									restartFromFailedRow = true;
									SkippingRowBasedOnCountryAndGroup(InvoicingData);
							    	BackButtonClick();
									continue;
								}
						}
						Thread.sleep(5000);
    					//Line Item 
						try{
						GridRowXpathValue = property.getProperty("LineItem");
						NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
    					new Select(driver.findElement(By.xpath(NewXpathValue))).selectByVisibleText(LineItem);  
    					PopupHandle();
    					SNOWLog.info("Line item entered successfully.Entered value from input excel sheet is: " + LineItem);
						}catch(Exception e)
						{
							try
							{
								GridRowXpathValue = property.getProperty("LineItem");
								NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
		    					new Select(driver.findElement(By.xpath(NewXpathValue))).selectByVisibleText(LineItem);  
		    					PopupHandle();
							}
							catch(Exception ee)
							{
							SNOWLog.info("Invalid Line Item Value");
							SNOWLog.info(LineItem);
							ExcelStatusWriter("Failure : Invalid Line Item",false);
							SkippingRowBasedOnCountryAndGroup(InvoicingData);
					    	BackButtonClick();
							continue;
							}
						}
						
						//waitForPageLoaded(driver);
				        
	    				//Text Line 
    					GridRowXpathValue = property.getProperty("TextLine");
    					NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
    					if(MaterialSalesText.equalsIgnoreCase("BLANK"))
    						EraseFieldData(NewXpathValue);
    					else if(driver.findElement(By.xpath(NewXpathValue)).isEnabled()&&!(MaterialSalesText.equalsIgnoreCase("NA")))
    					{
    						driver.findElement(By.xpath(NewXpathValue)).clear();
    						driver.findElement(By.xpath(NewXpathValue)).sendKeys(MaterialSalesText);
    						SNOWLog.info("Text Line item entered successfully.Entered value from input excel sheet is: " + MaterialSalesText);
    					}
    					
    					//Additional information - 529463 bug fixed 
    					try{
		    					GridRowXpathValue = property.getProperty("AdditionalInformation");
		    					NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
		    					if(AdditionalInformation.equalsIgnoreCase("BLANK"))
		    						EraseFieldData(NewXpathValue);
		    					else if(driver.findElement(By.xpath(NewXpathValue)).isEnabled()&&(!AdditionalInformation.equalsIgnoreCase("NA")))
		    					{
		    						driver.findElement(By.xpath(NewXpathValue)).clear();
		    						driver.findElement(By.xpath(NewXpathValue)).sendKeys(AdditionalInformation);
		    						SNOWLog.info("Additional information entered successfully.Entered value from input excel sheet is: " + AdditionalInformation);
		    					}
		    				}catch(Exception e){
    						SNOWLog.info("Additional information field is not present for the given GU and country.");
    					}
    					
    					//Quantity
    					GridRowXpathValue = property.getProperty("Quantity");
    					NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
						driver.findElement(By.xpath(NewXpathValue)).clear();
    					driver.findElement(By.xpath(NewXpathValue)).sendKeys(Quantity);
    					SNOWLog.info("Quantity entered successfully.Entered value from input excel sheet is: " + Quantity);
		        
    					//Amount
    					GridRowXpathValue = property.getProperty("Amount");
    					NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
    					driver.findElement(By.xpath(NewXpathValue)).clear();
    					driver.findElement(By.xpath(NewXpathValue)).sendKeys(Amount);
    					SNOWLog.info("amount entered sucessfully.Entered value from input excel sheet is: " + Amount);
			        	
	    				//Are you contractually obligated to bill reimbursable expenses as services revenue? (This is not common)
	    					GridRowXpathValue = property.getProperty("ContractuallyObligated");
	    					NewXpathValue = GridRowXpathValue.replaceAll("iGridRowIteratorval", chariGridRowIteratorval);
	    					if(!ContractualObligation.equalsIgnoreCase("NA"))
	    					{
		    					new Select(driver.findElement(By.xpath(NewXpathValue))).selectByVisibleText(ContractualObligation);
		    					SNOWLog.info("Are you contractually obligated input value from the excel sheet: " + ContractualObligation);
		    					if(PaymentTerms.equalsIgnoreCase("yes"))
		    						PopupHandle();
		    					PopupHandle();
	    					}
	    					SNOWLog.info("Entered Grid row information for the group: " + currentGroup  );
	    					
						if(IsUploadAttachmentRequired.equalsIgnoreCase("yes"))
						{
							//Attach the file for the current PO/Record
							if(!AttachmentPath.equalsIgnoreCase("NA"))
							{
								SNOWLog.info("The attachment file path from the input excel sheet is: "+AttachmentPath);
								if (!new File(AttachmentPath).exists())
								{
						        	SNOWLog.info("Timesheet not found at the given path.Check availability before re-running the bot.");
						        	ExcelStatusWriter("Failure : Attachment Not Found",true);
						        	restartFromFailedRow = true;
						        	SkippingRowBasedOnCountryAndGroup(InvoicingData);
							    	BackButtonClick();
									continue;
								}
						      //Attachment Button
						        clickExecute(driver, By.id(property.getProperty("AttachmentButton")));
						        Thread.sleep(500); 
						        SNOWLog.info("Upload Attachment button clicked successfully..");
						        
						    	//Attach File
						        Thread.sleep(500);
						        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='attachFile']")));
						        driver.findElement(By.xpath("//*[@id='attachFile']")).sendKeys(AttachmentPath);
						        
						    	SNOWLog.info("Started upload attachment process.");
						    	SNOWLog.info("filename : '"  + AttachmentPath + "'");
						
						    	String srcimage = driver.findElement(By.id("please_wait")).getAttribute("style"); 
								while(!(srcimage.contains("display: none")))
								{
									Thread.sleep(1000);
									SNOWLog.info("waiting for file to be upload..");
									srcimage = driver.findElement(By.id("please_wait")).getAttribute("style");
									continue;
								}
								SNOWLog.info("Completed upload attachment process..");
						    	clickExecute(driver, By.id(property.getProperty("AttachmentTab")));
						    	SNOWLog.info("came out of upload attachment window sucessfully.");
							}
							else
								SNOWLog.info("As Attachment path is not provided, will not attach file.");
						}
						//Logic for peeping in next PO to check if it is same as current
						//If check for iExcelRow and InvoicingData lenght to avoid array out of bound exception
						if(iExcelRow+1 <InvoicingData.length)
						{
							//Checking for the empty row
							if(!InvoicingData[iExcelRow+1].isEmpty())
							{
								int counter = iExcelRow+1;
								String ExcelRowInvoicingData[] = InvoicingData[iExcelRow+1].split("--");
								String countrycheck = ExcelRowInvoicingData[RPAStatusColumn];
								while(!(countrycheck.contains("Failure")||countrycheck.equalsIgnoreCase("Create Invoice"))&&(counter+1<InvoicingData.length))
								{
									counter++;
									SNOWLog.info("counter:"+counter);
									String	ExcelnextRowInvoicingData[] = InvoicingData[counter].split("--");
									countrycheck = ExcelnextRowInvoicingData[RPAStatusColumn];	
								}
								iExcelRow = counter-1;
								String	ExcelnextRowInvoicingData[] = InvoicingData[iExcelRow+1].split("--");
								Nextgroup = ExcelnextRowInvoicingData[GroupColumn];
								NextRowStatus=ExcelnextRowInvoicingData[RPAStatusColumn];
								SNOWLog.info("Next row group from the input excel sheet is: " + Nextgroup);
								SNOWLog.info("Next row status from the input excel sheet is: " + NextRowStatus);
								NextCountry = ExcelnextRowInvoicingData[CountryColumn];
								SNOWLog.info("The country for the next excel row is: "+NextCountry);
								//end of if excel row empty loop 
							}	
								//If current PO is not equals to next PO
								if((!currentGroup.equalsIgnoreCase(Nextgroup)))
								{
									//Logic for Saving Invoice and entering remaining details to the page
									Boolean flag = AddOtherDetails(InvoicingData);
									if(flag == false)
										continue;
									SNOWLog.info("Saving and creating the invoice for the group: " + currentGroup  );
							         Nextgroup="";
							         
							         /**
							          * TODO: Will need to check if IF condition is being used, if not code will be removed
							          */
//							         ExcelStatusWriter("Successfully generated the Invoice.",true, "No");
							         ExcelStatusWriter("Successfully generated the Invoice.",true);
							         
									//clicking on add invoice button if current and next country are same
							    	if(CurrentCountry.equalsIgnoreCase(NextCountry)&&(NextRowStatus.contains("Failure")||NextRowStatus.equalsIgnoreCase("Create Invoice")))
							    	{
							    		//TODO:
							    		cntFlag.add(iExcelRow);
							    		countryFlag = cntFlag.size();
							    									    									    		
							    		ExcelStatusWriter("Successfully added the PO and the country for the next invoice is same.Will Click on add invoice button.",false);
							    		ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RequestNumbercolumn, "Refer the last invoice template with same country for the request number. ");
							    		AddInvoiceButtonClick();							    		
							    		NextCountry="";
							    		NextRowStatus="";
							    		IsaddingInvoiceflag=true;
//								        if(!(Country.isEmpty()))
//								        {
//									        Thread.sleep(7000);
//									        PopupHandle();	
//								        }								        
							    		continue;
							    		
							    	}							    
							    	
							    	//checking whether user want to submit the the reports or not 
							    	if(WantToSubmitReport.equalsIgnoreCase("yes"))
						    		{
							    		SubmitInvoiceTicket();
						    		}
							    		
							    	else
							    		BackButtonClick();
								}
								else
								{
									//Before clicking first update the status for the current record as success
									if(RPAStatus.contains("Failure")||RPAStatus.equalsIgnoreCase("Create Invoice"))
										ExcelStatusWriter("Successfully added the template",false);
										ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RequestNumbercolumn, "Refer to the last template with same group for the Request number.");
									
									//Add grid row button click
									if(CurrentCountry.equalsIgnoreCase(NextCountry)&&(NextRowStatus.contains("Failure")||NextRowStatus.equalsIgnoreCase("Create Invoice")))
									{
										startRun = false;
										clickExecute(driver, By.id(property.getProperty("GridRowAdd")));
										SNOWLog.info("As the current and next row group are same so clicked on Add grid row button.");
										NextRowStatus="";
										Thread.sleep(500); //Initially it was 2000
									}
									else
									{
										startRun = false;
										log.info("The current and next country are not same so wont click add grid button.");
										Nextgroup="";
										Boolean flag = AddOtherDetails(InvoicingData);
										if(flag == false)
											continue;
								         SNOWLog.info("Saving Invoice:  " + currentGroup  );
								         NextCountry="";
								         
								         /**
								          * TODO: Will need to check if IF condition is being used, if not code will be removed
								          */
								         
								         //ExcelStatusWriter("Successfully added the template and generated the Invoice. ",true,"No");
								         ExcelStatusWriter("Successfully added the template and generated the Invoice. ",true);
								               
								    	//checking whether user want to submit the the reports or not 
								    	if(WantToSubmitReport.equalsIgnoreCase("yes"))
							    		{
								    		SubmitInvoiceTicket();
							    		}
								    	else
								    		BackButtonClick();
									}
							}
								//end if
								}
								else if(iExcelRow+1 ==InvoicingData.length)
								{
									//Logic for saving last row and saving other row information
									Boolean flag = AddOtherDetails(InvoicingData);
									if(flag == false)
										continue;						
							        SNOWLog.info("Saving Last Invoice:  " + currentGroup  );
							         NextCountry="";
							         
							         /**
							          * TODO: Will need to check if IF condition is being used, if not code will be removed
							          */
							         
							         //ExcelStatusWriter("Successfully added the template and generated the Invoice.",true,"No");
							         ExcelStatusWriter("Successfully added the template and generated the Invoice. ",true);
										
							         //clicking on add invoice button if current and next country are same
							    	if(CurrentCountry.equalsIgnoreCase(NextCountry)&&(NextRowStatus.contains("Failure")||NextRowStatus.equalsIgnoreCase("Create Invoice")))
							    	{
							    		/**
								          * TODO: Will need to check if IF condition is being used, if not code will be removed
								          */
							    		
							    		//TODO:
							    		cntFlag.add(iExcelRow);
							    		countryFlag = cntFlag.size();
							    		
							    		//Status log for the current record
							    		//ExcelStatusWriter("Successfully added the template and the country for the next invoice is same.Will Click on add invoice button.",true,"No");
							    		ExcelStatusWriter("Successfully added the template and the country for the next invoice is same.Will Click on add invoice button.",true);
							    		AddInvoiceButtonClick();
							    		NextCountry="";
							    		NextRowStatus="";
							    		IsaddingInvoiceflag=true;
//								        if(!(Country.isEmpty()))
//								        {
//									        Thread.sleep(7000);
//									        PopupHandle();	
//								        }
							    		continue;

							    	}							   							    	
							    	
							    	//checking whether user want to submit the the reports or not 
							    	if(WantToSubmitReport.equalsIgnoreCase("yes"))
							    		SubmitInvoiceTicket();
							    	else
							    		BackButtonClick();
								}
					}
					else
						SNOWLog.info("The status for the current row is other than 'Create Invoice' or 'Failure'.So skipping the record and moving to the next one.");
					} //Close If for handling delivery method error
					
				}//If close to check excel InvoicingData
			}//For loop close
			//After completion of all data insertion sending success report
			Runtime.getRuntime().exec(new String[] {"wscript.exe","C:\\PCS RPA\\DigitalInvoicing\\ServerFiles\\" + "SendGenerateInvoicesSuccessEmail.vbs"});
			}catch(Exception e)
		{
			if(flag==true)
			{
				SNOWLog.info("Successfully Completed the automation,Bot will exit now.");
				Runtime.getRuntime().exec(new String[] {"wscript.exe","C:\\PCS RPA\\DigitalInvoicing\\ServerFiles\\" + "SendGenerateInvoicesSuccessEmail.vbs"});
				Thread.sleep(5000);
			}
			else{
				SNOWLog.info("Error occured while data insertion into the application.Restarting the automation.");
				SNOWLog.info(e.getLocalizedMessage());
				ExcelStatusWriter("Failure : Failed to create the Invoice",true);
				SkippingRowBasedOnCountryAndGroup(InvoicingData);
//				ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, 0, 51,Integer.toString(iExcelRow));//written AL cell 
				Runtime.getRuntime().exec(new String[] {"wscript.exe","C:\\PCS RPA\\DigitalInvoicing\\ServerFiles\\" + "SendGenerateInvoicesErrorMail.vbs"});
				
				File dir = new File("C:\\PCS RPA\\DigitalInvoicing\\BOTs");
				ProcessBuilder pb;
				pb = new ProcessBuilder("cmd.exe","/c","start","/w","cmd", "/c", "LaunchBOT_RaiseSNOWInvoiceRequest_v1.0.bat");
				pb.directory(dir);
				pb.start();
				SNOWLog.info("Restarted RPA..");
			}
		}
}
	
	public Boolean AddOtherDetails(String InvoicingData[]) throws Exception
	{
			//Payment terms
			IsDropdownfieldvalueFound = InsertMandatorydropDownFieldData("Payment terms",PaymentTerms,"PaymentTerms",InvoicingData);
			if(IsDropdownfieldvalueFound == false)
				return false;
			if(PaymentTerms.equals("Z000 - Due Immediately"))
				PopupHandle();
			PopupHandle();
				
				Thread.sleep(2000);
				
				//Payment method
				InsertMandatorydropDownFieldData("Payment method",PaymentMethod,"PaymentMethod",InvoicingData);
				if(IsDropdownfieldvalueFound == false)
					return false;
				PopupHandle();
				
				Boolean IsEnabled;
				
				//Add Value Days
				String check = (driver.findElement(By.id("IO:e1467e2ddb408b00f945f9a41d9619d2")).getAttribute("class"));
				
				if((check.contains("disabled"))||(check.contains("readonly")))
				{
					IsEnabled = false;
				}
				else 
				{
					IsEnabled = true;
				}
				
				
//				String IsEnabled = driver.findElement(By.id(property.getProperty("AddValueDays"))).getAttribute("readonly");
				
				SNOWLog.info("Is add value days field editable = "+IsEnabled);
				
				if(IsEnabled==true&&(! AddValueDays.equalsIgnoreCase("NA")))
				{
					EraseFieldData("AddValueDays");
					AddValueDays=AddValueDays.replace(".0", "");
					InsertTextFieldData("Add value days",AddValueDays,"AddValueDays");
				}
				PopupHandle();
			
				//Fixed value date
//				IsEnabled = driver.findElement(By.id(property.getProperty("FixedValueDays"))).getAttribute("readonly");
				
				String check2 = (driver.findElement(By.id("IO:e1467e2ddb408b00f945f9a41d9619d2")).getAttribute("class"));
				
				if((check2.contains("disabled"))||(check.contains("readonly")))
				{
					IsEnabled = false;
				}
				else 
				{
					IsEnabled = true;
				}
				
				SNOWLog.info("Is Fixed value date field editable = "+IsEnabled);
				
				if(IsEnabled=true&&(!FixedValueDate.equalsIgnoreCase("NA")))
					InsertTextFieldData("Fixed value date",FixedValueDate,"FixedValueDays");
				PopupHandle();

		//Comments
		if(Comments.equalsIgnoreCase("BLANK"))
			EraseFieldData("Comments");
		else if(!Comments.equalsIgnoreCase("NA"))
		{
			String comment = driver.findElement(By.id("IO:ef2376a9db408b00f945f9a41d961942")).getText();
			
			if(comment.contains(WebPortal))
				driver.findElement(By.id("IO:ef2376a9db408b00f945f9a41d961942")).sendKeys("\n"+Comments);
			else {
				driver.findElement(By.id("IO:ef2376a9db408b00f945f9a41d961942")).sendKeys(Comments);
			}
			SNOWLog.info(Comments+" entered successfully.Entered value from input excel sheet is: "+ Comments);
		}
		
		//Form Header
		if(FormHeader.equalsIgnoreCase("BLANK"))
			EraseFieldData("FormHeader");
		else if(!FormHeader.equalsIgnoreCase("NA"))
			InsertTextFieldData("Form Header",FormHeader,"FormHeader");
		
		//Additional Note 1
		if(AdditionalNote1.equalsIgnoreCase("BLANK"))
				EraseFieldData("AdditionalNote1");
		else if(!AdditionalNote1.equalsIgnoreCase("NA"))
				InsertTextFieldData("Additional Note 1",AdditionalNote1,"AdditionalNote1");
		
		//Additional Note 2
		if(AdditionalNote2.equalsIgnoreCase("BLANK"))
			EraseFieldData("AdditionalNote2");
		else if(!AdditionalNote2.equalsIgnoreCase("NA"))
			InsertTextFieldData("Additional Note 2",AdditionalNote2,"AdditionalNote2");
		    
		//Customer additional info
			if(CustomerAddInfo.equalsIgnoreCase("BLANK"))
				EraseFieldData("CustomerAddInfo");
			else if(!CustomerAddInfo.equalsIgnoreCase("NA"))
				InsertTextFieldData("Customer Additional infomation",CustomerAddInfo,"CustomerAddInfo");
	    
		//Company additional info
			if(CompanyAddInfo.equalsIgnoreCase("BLANK"))
				EraseFieldData("ComapanyAddInfo");
			else if(!CompanyAddInfo.equalsIgnoreCase("NA"))
				InsertTextFieldData("Company additional infomation",CompanyAddInfo,"ComapanyAddInfo");
		
		//Bill-to-Party
			if(BillToParty.equalsIgnoreCase("BLANK"))
				EraseFieldData("Bill-to-Party");
			else if(!BillToParty.equalsIgnoreCase("NA"))
			{
				driver.findElement(By.id("IO:cb89feaddb408b00f945f9a41d961959")).clear();
				driver.findElement(By.id("IO:cb89feaddb408b00f945f9a41d961959")).sendKeys(BillToParty);
				SNOWLog.info(BillToParty+" entered successfully.Entered value from input excel sheet is: "+ BillToParty);
			}
		
		//Logo
			if(!Logo.equalsIgnoreCase("NA"))
			{
				
				InsertOptionaldropdownFieldData("Logo",Logo,"Logo");
			}
			
		//Layout		
			InsertOptionaldropdownFieldData("Layout",Layout,"Layout");
			
		//Should Billing Team Deliver Invoice
			InsertOptionaldropdownFieldData("Should billing team deliver invoice",ShouldBillingTeamDeliverInvoice,"BillingTeamDeliverInvoice");
			
				if (ShouldBillingTeamDeliverInvoice.equalsIgnoreCase("Yes"))
				{
					//Delivery method
					if(DeliveryMethod.equalsIgnoreCase("None"))
					{
						SNOWLog.info("As the delivery method given as none..");
						InsertOptionaldropdownFieldData("Delivery Method","-- None --","DeliveryMethod");
					}
					else if(DeliveryMethod.equals("E-mail")||(DeliveryMethod.equals("E-mail + Post")))
					{
						InsertOptionaldropdownFieldData("Delivery Method",DeliveryMethod,"DeliveryMethod");
						InsertTextFieldData("Email",Email,"Email"); 
					}
					else if(DeliveryMethod.equals("Post"))
						
					{
						InsertOptionaldropdownFieldData("Delivery Method",DeliveryMethod,"DeliveryMethod");
						//Postal Address
						driver.findElement(By.id("IO:93ea7eeddb408b00f945f9a41d961973")).clear();
						driver.findElement(By.id("IO:93ea7eeddb408b00f945f9a41d961973")).sendKeys(PostalAddress);
						SNOWLog.info(PostalAddress+" entered successfully.Entered value from input excel sheet is: "+ PostalAddress);
					}
					else if(DeliveryMethod.equalsIgnoreCase("IDOC"))
					{
						InsertOptionaldropdownFieldData("Delivery Method",DeliveryMethod,"DeliveryMethod");
						//If the delivery method is IDOC for Norway,Sweden,Denmark,Finland
						PopupHandle();
						PopupHandle();
						if(IdentificationNumber.contains(".0"))
							IdentificationNumber = IdentificationNumber.replace(".0","");
						InsertTextFieldData("Identification Number",IdentificationNumber,"IdentificationNumber");
					}
					else if(DeliveryMethod.contains("Web Portal")) 
					{					
						InsertOptionaldropdownFieldData("Delivery Method",DeliveryMethod,"DeliveryMethod");
						
						try{
							new Select(driver.findElement(By.id("IO:275df261db808b00f945f9a41d96199c"))).selectByVisibleText(WebPortal);
							SNOWLog.info("Web Portal is entered sucessfully.Entered value from input excel sheet is: "+WebPortal);
						}catch(Exception e)
						{
							InsertOptionaldropdownFieldData("Web Portal","Other- Please specify in Comments column","WebPortal");
							driver.findElement(By.id("IO:ef2376a9db408b00f945f9a41d961942")).sendKeys("\nOther Web Portal: " +WebPortal);
							SNOWLog.info(WebPortal+" entry is not available. Other- Please specify in Comments column drop down option is selected.");
							SNOWLog.info("Other web portal is successfully entered : "+WebPortal);
						}
						
					}
					else if(DeliveryMethod.contains("Web Portal + Post")) 
					{
						InsertOptionaldropdownFieldData("Delivery Method",DeliveryMethod,"DeliveryMethod");
						
						try{
							new Select(driver.findElement(By.id("IO:275df261db808b00f945f9a41d96199c"))).selectByVisibleText(WebPortal);
							SNOWLog.info("Web Portal is entered sucessfully.Entered value from input excel sheet is: "+WebPortal);
						}catch(Exception e)
						{
							InsertOptionaldropdownFieldData("Web Portal","Other- Please specify in Comments column","WebPortal");
							driver.findElement(By.id("IO:ef2376a9db408b00f945f9a41d961942")).sendKeys("\nOther Web Portal: " +WebPortal);
							SNOWLog.info(WebPortal+" entry is not available. Other- Please specify in Comments column drop down option is selected.");
							SNOWLog.info("Other web portal is successfully entered : "+WebPortal);
						}
						
						//Postal Address
						driver.findElement(By.id("IO:93ea7eeddb408b00f945f9a41d961973")).clear();
						driver.findElement(By.id("IO:93ea7eeddb408b00f945f9a41d961973")).sendKeys(PostalAddress);
						SNOWLog.info(PostalAddress+" entered successfully.Entered value from input excel sheet is: "+ PostalAddress);
					}				
					else
						InsertOptionaldropdownFieldData("Delivery Method",DeliveryMethod,"DeliveryMethod");
						
					//Cover Letter needed? 
					InsertOptionaldropdownFieldData("Cover Letter needed",CoverLetterNeeded,"CoverLetterNeeded");					
						
					//Include Attachment mandatory 			
					if(IncludeAttachment.equalsIgnoreCase("No")) 
					{
						InsertOptionaldropdownFieldData("Include Attachment",IncludeAttachment,"IncludeAttachment");
					}
					else
					{
						InsertOptionaldropdownFieldData("Include Attachment",IncludeAttachment,"IncludeAttachment");
						
						//Attachment Required. Take note of this after adding NA in Test Data
						if(AttachmentRequired.equalsIgnoreCase("BLANK"))
							EraseFieldData("AttachmentRequired");
						else if(!AttachmentRequired.equalsIgnoreCase("NA"))
							InsertOptionaldropdownFieldData("Attachment Required",AttachmentRequired,"AttachmentRequired");
					}
				}
			
			//SupplementalDocument
				InsertOptionaldropdownFieldData("Supplemental Document",SupplementalDocument,"SupplementalDocument");
				
				try{
						wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='dsm_dmr_regular_confirm_fa']//button[@type='submit']")));
						driver.findElement(By.xpath("//*[@id='dsm_dmr_regular_confirm_fa']//button[@type='submit']")).click();
						SNOWLog.info("Clicked on supplemental document confirmation alert successfully.");
				}catch(Exception e)
				{
					SNOWLog.info("Supplemental document alert is not present.");
				}
					
			//Please Provide Explanation
				InsertTextFieldData("Please Provide Explanation",PleaseProvideExplanation,"PleaseProvideExplanation");
		
	        //Reviewer ID
				if(!ReviewerId.equalsIgnoreCase("NA"))
				{
					driver.findElement(By.id(property.getProperty("ReviewerIDUnlockButton"))).click();
					SNOWLog.info("Clicked on Reviewer Id Unlock Icon successfully.");
					InsertTextFieldData("Reviewer ID ",ReviewerId,"ReviewerIDTextArea");
					Thread.sleep(3000);
					hitEnterKey();
				}

			//Approver ID
				if(!ApproverId.equalsIgnoreCase("NA"))
				{
					SNOWLog.info("Approver IDs considered for updation from the input excel sheet: "+ApproverId);
					ApproverIDInsertion();
				}
				
		    	//TODO: Catching invoice number
				driver.switchTo().defaultContent();
				
		    	invoiceCntr = 1;
				Boolean compare = true;
													
				while(compare==true)
				{
					try {
					Boolean chck = (driver.findElement(By.xpath("//span[@class='panel-body']//div[4]//span[@class='left-margin ng-scope']["+invoiceCntr+"]")).isDisplayed());	
//					System.out.println(chck);
					compare = chck;
					invoiceCntr++;
					}
					catch (Exception e) {
						compare = false;
					}
				}
			
				print = invoiceCntr-1;
//				System.out.println(print);
				
				return true;
         
	}
	
	public void InsertTextFieldData(String FieldName,String InputValue,String xpath)
	{
		try{
		driver.findElement(By.id(property.getProperty(xpath))).clear();
		driver.findElement(By.id(property.getProperty(xpath))).sendKeys( InputValue);
		SNOWLog.info(FieldName+" entered successfully.Entered value from input excel sheet is: "+ InputValue);
		}catch(Exception e)
		{
			SNOWLog.info(FieldName+" field not present.");
		}
	}
	
	public boolean InsertMandatorydropDownFieldData(String FieldName,String InputValue,String Xpath,String InvoicingData[] ) throws Exception
	{
		 try{
		        new Select(driver.findElement(By.id(property.getProperty(Xpath)))).selectByVisibleText(InputValue);
		        SNOWLog.info(FieldName + " entered successfully.Entered value is: "+InputValue); 
		        }catch(Exception e)
		        {
		        	SNOWLog.info(FieldName+" dropdown value is not found.Will skip this and move ahead.");
					ExcelStatusWriter("Failure : "+FieldName+" field value is not found.",true);
					restartFromFailedRow = true;
					SkippingRowBasedOnCountryAndGroup(InvoicingData);
			    	BackButtonClick();
					return false;
		        }
		return true;
	}
	
	public void InsertOptionaldropdownFieldData(String FieldName,String InputValue,String Xpath)
	{
		try{
			new Select(driver.findElement(By.id(property.getProperty(Xpath)))).selectByVisibleText(InputValue);
			SNOWLog.info(FieldName+" entered sucessfully.Entered value from input excel sheet is: "+InputValue);
		}catch(Exception e)
		{
			SNOWLog.info(FieldName+" dropdown field or value is not present.Value tried to enter was: "+InputValue);
		}
	}
	
	public void SkippingRowBasedOnCountryAndGroup(String InvoicingData[]) throws Exception
	{
		//Checking and skipping the next invoices with the same country and same group
		if(iExcelRow+1<InvoicingData.length)
		{
			int counter = iExcelRow+1;
			String ExcelRowInvoicingData[] = InvoicingData[iExcelRow+1].split("--");
			String NextCountryCheck = ExcelRowInvoicingData[CountryColumn];
			String NextGroupCheck = ExcelRowInvoicingData[GroupColumn];
			
			/**TODO:
			 * 
			 */
			
			String ExcelCurRowInvoicingData[] = InvoicingData[iExcelRow].split("--");
			String StatusCheck = ExcelCurRowInvoicingData[RPAStatusColumn];
			
			System.out.println(StatusCheck);
			
			//if ((StatusCheck.toLowerCase().contains("not found"))&&(CurrentCountry.equalsIgnoreCase(NextCountryCheck)||currentGroup.equalsIgnoreCase(NextGroupCheck)))
			if ((restartFromFailedRow==true)&&(CurrentCountry.equalsIgnoreCase(NextCountryCheck)||currentGroup.equalsIgnoreCase(NextGroupCheck)))
			{
				IsaddingInvoiceflag=false;
			}
			else
			{
				while((CurrentCountry.equalsIgnoreCase(NextCountryCheck)||currentGroup.equalsIgnoreCase(NextGroupCheck))&&counter+1<=InvoicingData.length)
				{
					SNOWLog.info("Row number considered for checking: "+counter+" With current row country: "+CurrentCountry+" Next Country: "+NextCountryCheck+" Current group: "+currentGroup+" Next group: "+NextGroupCheck);
					SNOWLog.info("Failure : As country or group is similar so skipping");
					ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, counter, RPAStatusColumn, "Failure : As country or group is similar so skipping");
					try{
						counter++;
						String	ExcelnextRowInvoicingData[] = InvoicingData[counter].split("--");
						NextCountryCheck = ExcelnextRowInvoicingData[CountryColumn];
						NextGroupCheck =ExcelnextRowInvoicingData[GroupColumn];
//						System.out.println(NextGroupCheck);
					}catch(Exception e1)
					{
						IsProcessingLastRecordRow = true;
						flag=true;
						IsaddingInvoiceflag=false;
						SNOWLog.info("Last record for processing.");
					}		
				}
				if(IsProcessingLastRecordRow==false)
				{
					iExcelRow = counter-1;
					IsaddingInvoiceflag=false;
				}else
				{
					SNOWLog.info("As all records were processed so will exit.");
				}	
			}
			
		}
	}
	
	public void EraseFieldData(String fieldName) //If user wants to clear up the data and he provide BLANK as input
	{		
		try{
				driver.findElement(By.id(property.getProperty(fieldName))).clear();
				SNOWLog.info("Blanked out the field data.");
		}catch(Exception e)
		{
			SNOWLog.info("The field does not present: "+fieldName);
		}
	}
	
	public void ApproverIDInsertion() throws InterruptedException, AWTException
	{
		try
		{
			driver.findElement(By.id(property.getProperty("ApproverIDUnlockButton"))).click();
			SNOWLog.info("Clicked on Reviewer Id Unlock Icon successfully.");
		}
		
		catch(Exception e)
		{
			SNOWLog.info("Approver ID unlock Icon clicked successfully.");
		}
		
        if(ApproverId.contains(";"))
        {
        	String Approver[] = ApproverId.split(";");
        	
        	for(int i=0;i<Approver.length;i++)
        	{
        		driver.findElement(By.id(property.getProperty("ApproverIDTextArea"))).sendKeys(Approver[i]);
        		Thread.sleep(3000);
        		hitEnterKey();
        		SNOWLog.info("Approver ID Added successfully.Entered value is: "+Approver[i]);   	
        	}	
        }
        else
        {
        	driver.findElement(By.id(property.getProperty("ApproverIDTextArea"))).sendKeys(ApproverId);
			Thread.sleep(3000);
			hitEnterKey();
    		SNOWLog.info("Approver ID Added successfully.Entered value is: "+ApproverId);  
        }
        
	}

	public void AddInvoiceButtonClick() throws InterruptedException
	{
		negTest=testFlag;
		startRun = false;
		
		SNOWLog.info("As The current and next country are same so will click on add invoice button.");
		driver.switchTo().defaultContent();
		//add invoice button
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(property.getProperty("AddInvoice"))));
		driver.findElement(By.xpath(property.getProperty("AddInvoice"))).click();
		SNOWLog.info("clicked on add Invoice button successfully.");

		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(property.getProperty("Add"))));
		driver.findElement(By.xpath(property.getProperty("Add"))).click();
		SNOWLog.info("clicked on add button successfully.");
		
		String InvoiceEditXpath= "//span["+InvoiceNumber+"]/span/a[text()='Invoice']";
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(InvoiceEditXpath)));
		driver.findElement(By.xpath(InvoiceEditXpath)).click();
		SNOWLog.info("Clicked on the invoice for editing...");
		InvoiceNumber++;
	}
	
	public void BackButtonClick() throws Exception
	{
		driver.switchTo().defaultContent();
		IsaddingInvoiceflag=false;
		clickExecute(driver, By.xpath(property.getProperty("BackButton")));
		SNOWLog.info("Clicked Back Button successfully..");
		Thread.sleep(5000);//Necessary for coming back to the home page
		SNOWLog.info("came out of back button sleep....");
		
		int backCnt = iExcelRow;
		
		for(int back=1;back<=10;back++)
		{
			String DataBack[] = ExcelFunctions.readData(InputDataFile, InputSheetname);
			String CurRowBackData[] = DataBack[backCnt].split("--");
			String PrevRowBackData[] = DataBack[backCnt-1].split("--");
			String PrevCCGroupCheck = PrevRowBackData[CountryColumn];
			String CurrentCCGroupCheck = CurRowBackData[CountryColumn];
			
    		if(CurrentCCGroupCheck.equalsIgnoreCase(PrevCCGroupCheck))
			{
    			backCnt--;
				SNOWLog.info("Error encountered on last invoice.");
				ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, backCnt, RPAStatusColumn, "Error encountered on last invoice.");
			}
    		else
    		{
    			break;
    		}
		}
	}
	
	public void SubmitInvoiceTicket() throws Exception
	{
//			System.out.println(print);
			
			EndTime = System.currentTimeMillis();
			TotalTimeTaken = EndTime - startTime;
			TotalTimeTaken = TotalTimeTaken/1000;
			TotalTimeTaken = TotalTimeTaken/60;
			String TimeTaken = Double.toString(TotalTimeTaken);
		
			driver.switchTo().defaultContent();
			clickExecute(driver, By.xpath(property.getProperty("SubmitAllButton")));
	
			wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@class='modal-dialog ']")));
			String SubmitMessage = driver.findElement(By.xpath("//div[@class='modal-dialog ']//*[contains(@ng-if,'options.message')]")).getText();
			SNOWLog.info("On Clicking submit button message:"+SubmitMessage);
			
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='modal-dialog ']//button[contains(text(),'Yes')]")));
			driver.findElement(By.xpath("//div[@class='modal-dialog ']//button[contains(text(),'Yes')]")).click(); 
			SNOWLog.info("Submit All invoice button clicked.");
			
			Boolean cont = true; 
					
			try
			{
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[@id='uiNotificationContainer']//span")));
			}
			catch(Exception e)
			{
		    	ExcelStatusWriter("Failure:Error occured while submitting the invoice.Check for all mandatory fields.", true);
				BackButtonClick();
				cont = false;
			}
			
			
			
//			try{	 
			
			if(cont==true)
			{

				String RequestID = driver.findElement(By.xpath("//div[@id='uiNotificationContainer']//span")).getText();
				
				SNOWLog.info("Submitted the Invoice successfully and  REQ number generated..Please check your input file for the same");
		    	SNOWLog.info("Request id generated after submitting the request:"+RequestID);
		    	ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RequestNumbercolumn, RequestID);
	    	
		    	/**
		    	 * TODO : CAPTURE RITM FOR SAME COUNTRY
		    	 */
		    		
		    	Thread.sleep(10000);
		    				    	
		    	try 
		    	{	
		    		wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//table[contains(@class,'table-responsive')]")));
		    		clickExecute(driver, By.xpath("//div[contains(text(),'Created')]"));
		    		clickExecute(driver, By.xpath("//div[contains(text(),'Created')]"));
		    		wait.until(ExpectedConditions.textToBe((By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[3]")), "Submitted/New Request"));
		    	}
		    	catch (Exception e)
		    	{
		    		SNOWLog.info("RITM and REQUEST NUMBER Capture Timeout: 60s");
		    		
			    	try 
			    	{	
			    		wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//table[contains(@class,'table-responsive')]")));
			    		clickExecute(driver, By.xpath("//div[contains(text(),'Created')]"));
			    		clickExecute(driver, By.xpath("//div[contains(text(),'Created')]"));
			    		wait.until(ExpectedConditions.textToBe((By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[3]")), "Submitted/New Request"));
			    	}
			    	catch (Exception e2)
			    	{
			    		SNOWLog.info("RITM and REQUEST NUMBER Capture Timeout: 120s");
			    		
				    	try 
				    	{	
				    		wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//table[contains(@class,'table-responsive')]")));
				    		clickExecute(driver, By.xpath("//div[contains(text(),'Created')]"));
				    		clickExecute(driver, By.xpath("//div[contains(text(),'Created')]"));
				    		wait.until(ExpectedConditions.textToBe((By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[3]")), "Submitted/New Request"));
				    	}
				    	catch (Exception e3)
				    	{
				    		SNOWLog.info("RITM and REQUEST NUMBER Capture Timeout: 180s");
				    	}
			    	}
		    	}
		    	
		    	int printCntr = 1;
		    	int printHalf = test/2;
		    
		    	if(print<11)
		    	{
		    		for(printCntr=1; printCntr<test; --printCntr)
		    		{
		    			try
		    			{
		    				String reqCheckLast = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')]["+print+"]/td[2]")).getText();
		    				String reqCheckFirst = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[2]")).getText();
				    		
				    		if(reqCheckFirst.equalsIgnoreCase(reqCheckLast))
				    		{
				    			SNOWLog.info("REQUEST NUMBER GENERATED: "+reqCheckLast);
				    			break;
				    		}
		    			}
		    			catch(Exception e)
		    			{
		    				Thread.sleep(3000);
		    				
		    				if(printCntr==printHalf)
				    		{
					    		driver.get(driver.getCurrentUrl());
					    		Thread.sleep(5000);
					    		SNOWLog.info("PAGE REFRESHED");
					    		printCntr++;
				    		}
		    			}
		    		}
		    		
		    		/**
		    		 * FOR RITM WRITING
		    		 */
		    		
		    		if(print>1)
			    	{
			    		
				    	int countryCnt = 1;
				    	int rowCnt = iExcelRow;
				    	
			    		while(countryCnt<=print)
			    		{
			    			String RITMNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')]["+countryCnt+"]/td[1]")).getText();
			    			String REQNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')]["+countryCnt+"]/td[2]")).getText();
			    			
			    			String reqCheckLast = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')]["+print+"]/td[2]")).getText();
		    				String reqCheckFirst = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')]["+countryCnt+"]/td[2]")).getText();
					    	
					    	if(reqCheckFirst.equalsIgnoreCase(reqCheckLast))
							{
					    		SNOWLog.info("RITM NUMBER GENERATED: "+RITMNumber);
					    		SNOWLog.info("REQUEST NUMBER GENERATED: "+REQNumber);
					    		
								ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt, RITMNumbercolumn, RITMNumber);
								ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt, RequestNumbercolumn, REQNumber);
								
								int i;
						    	for(i=1; 1<=10; i++)
						    	{
							    	String Data[] = ExcelFunctions.readData(InputDataFile, InputSheetname);
									String CurRowInvData[] = Data[rowCnt].split("--");
									String PrevRowInvData[] = Data[rowCnt-1].split("--");
									String PrevGroupCheck = PrevRowInvData[GroupColumn];
									String CurrentGroupCheck = CurRowInvData[GroupColumn];
									String PrevCountryCheck = PrevRowInvData[CountryColumn];
									String CurrentCountryCheck = CurRowInvData[CountryColumn];
									String StatusColumn[] = Data[rowCnt-1].split("--");
									String StatusCheck = StatusColumn[RPAStatusColumn];
									
						    		if((StatusCheck.toLowerCase().contains("success"))&&(CurrentGroupCheck.equalsIgnoreCase(PrevGroupCheck))&&(PrevCountryCheck.equalsIgnoreCase(CurrentCountryCheck)))
									{
										rowCnt--;
										SNOWLog.info("RITM NUMBER GENERATED: "+RITMNumber);
							    		SNOWLog.info("REQUEST NUMBER GENERATED: "+REQNumber);
										ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt, RITMNumbercolumn, RITMNumber);
										ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt, RequestNumbercolumn, REQNumber);
										ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt,TimeTakenToCreateInvoice,TimeTaken);
									}
						    		else
						    		{
						    			break;
			
						    		}
						    	}
								
					    		rowCnt--;
							}
					    					    	
					    	countryCnt++;
			    		}
			    		
			    	}
		    		
		    		else
			    	{		
		    			String RITMNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[1]")).getText();
		    			String REQNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[2]")).getText();
		    		   	
						SNOWLog.info("RITM NUMBER GENERATED: "+RITMNumber);
						SNOWLog.info("REQUEST NUMBER GENERATED: "+REQNumber);
						
				    	ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RITMNumbercolumn, RITMNumber);
				    	ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RequestNumbercolumn, REQNumber);
				    	
						int x=1;
						int rowCnt2 = iExcelRow;
						
						for(x=1;x<=10;x++)
						{
					    	String Data[] = ExcelFunctions.readData(InputDataFile, InputSheetname);
							String CurRowInvData[] = Data[rowCnt2].split("--");
							String PrevRowInvData[] = Data[rowCnt2-1].split("--");
							String PrevGroupCheck = PrevRowInvData[GroupColumn];
							String CurrentGroupCheck = CurRowInvData[GroupColumn];
							String PrevCountryCheck = PrevRowInvData[CountryColumn];
							String CurrentCountryCheck = CurRowInvData[CountryColumn];
							
							if((CurrentGroupCheck.equalsIgnoreCase(PrevGroupCheck))&&(PrevCountryCheck.equalsIgnoreCase(CurrentCountryCheck)))
							{
								rowCnt2--;
								SNOWLog.info("RITM NUMBER GENERATED: "+RITMNumber);
								SNOWLog.info("REQUEST NUMBER GENERATED: "+REQNumber);
								ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt2, RITMNumbercolumn, RITMNumber);
								ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt2, RequestNumbercolumn, REQNumber);
								ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt2,TimeTakenToCreateInvoice,TimeTaken);
							}
							else
								break;
						}
			    	}
		    		
		    	}
		    	else
		    	{
		    		SNOWLog.info("RITM Check limited to 10 invoice only.");
	    			
		    		for(printCntr=1; printCntr<test; --printCntr)
		    		{
		    			try
		    			{
		    				String reqCheckLast = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][10]/td[2]")).getText();
		    				String reqCheckFirst = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][1]/td[2]")).getText();

		    				String RITMNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][10]/td[1]")).getText();
			    			String REQNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][10]/td[2]")).getText();
		    				
				    		if(reqCheckFirst.equalsIgnoreCase(reqCheckLast))
				    		{
								SNOWLog.info("RITM NUMBER GENERATED: "+RITMNumber);
								SNOWLog.info("REQUEST NUMBER GENERATED: "+REQNumber);
				    			break;
				    		}
		    			}
		    			catch(Exception e)
		    			{
		    				Thread.sleep(3000);
		    				
		    				if(printCntr==printHalf)
				    		{
					    		driver.get(driver.getCurrentUrl());
					    		Thread.sleep(5000);
					    		SNOWLog.info("PAGE REFRESHED");
					    		printCntr++;
				    		}
		    			}
		    		}
		    		
			    	int rowCnt2 = iExcelRow;
			    						
    				String RITMNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][10]/td[1]")).getText();
	    			String REQNumber = driver.findElement(By.xpath("//tr[contains(@ng-repeat,'item in data.list')][10]/td[2]")).getText();
			    	
			    	ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt2, RITMNumbercolumn, RITMNumber);
			    	ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt2, RequestNumbercolumn, REQNumber);
					
			    	SNOWLog.info("LATEST RITM NUMBER GENERATED: "+RITMNumber);
					SNOWLog.info("LATEST REQUEST NUMBER GENERATED: "+REQNumber);
			    	
		    		int i;
			    	for(i=print; i>0; i--)
			    	{
				    	String Data[] = ExcelFunctions.readData(InputDataFile, InputSheetname);
						String CurRowInvData[] = Data[i].split("--");
						String PrevRowInvData[] = Data[i-1].split("--");
						String PrevCCCheck = PrevRowInvData[CountryColumn];
						String CurCCCheck = CurRowInvData[CountryColumn];
						
			    		if(CurCCCheck.equalsIgnoreCase(PrevCCCheck))
						{
			    			rowCnt2--;
							ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, rowCnt2, RITMNumbercolumn, "Please check generated RITM manually. Refer to latest REQ/RITM generated for the Request Number");
						}
			    		else
			    		{
			    			break;

			    		}
			    	}
		    	}
		    	
	    	
	    	driver.navigate().to(SNInvoicingURL);
	    	startRun = true;
			}
			
			
//			}catch(Exception e)
//			{
//		    	ExcelStatusWriter("Failure:Error occured while submitting the invoice.Check for all mandatory fields.", true);
//				BackButtonClick();
//			}
	    	IsaddingInvoiceflag=false;
	    	Thread.sleep(5000);
//	    	driver.navigate().to(“https://accentureinternalsbox6.service-now.com/dsm/?id=gw_dim_invoicing_entry_fa”);

	}
	
	public void hitEnterKey() throws AWTException {
			Robot robot = new Robot();
			try {
		    robot = new Robot();
			} 
			catch (AWTException e) {
				e.printStackTrace();
		}
		robot.keyPress(KeyEvent.VK_ENTER);
	}
	
	public void PopupHandle() throws InterruptedException
	{		
		//pop up  click
		int counter=0;
		while(counter++<10)
		{			
	        try
	        {
	            driver.switchTo().alert();
	            SNOWLog.info("Switch to alert sucessfully..");
	            SNOWLog.info("ALERT Message : "+driver.switchTo().alert().getText());
	            driver.switchTo().alert().accept();
	            SNOWLog.info("Clicked on alert sucessfully..");
	            break;
	        }
	        catch(NoAlertPresentException e)
	        {
	          Thread.sleep(1000);
	          continue;
	        }
	   }

	}
	
	public boolean isAlertPresent() {
	    try {
	        WebDriverWait wait = new WebDriverWait(driver, 1000);
	        wait.until(ExpectedConditions.alertIsPresent());
	        return true;
	    } // try
	    catch (Exception e) {
	        return false;
	    } // catch
	}
	
	public void ExcelStatusWriter(String Meassage,Boolean IsTimeToWrite) throws Exception
	{
		ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RPAStatusColumn, Meassage);
		if(IsTimeToWrite.equals(true))
		{
			EndTime = System.currentTimeMillis();
			SNOWLog.info("End time taken to complete the invoice process: "+EndTime);
			TotalTimeTaken = EndTime - startTime;
			TotalTimeTaken = TotalTimeTaken/1000;
			TotalTimeTaken = TotalTimeTaken/60;
			String TimeTaken = Double.toString(TotalTimeTaken);
			ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow,TimeTakenToCreateInvoice,TimeTaken);
		}
//		if(!(RequestNumber.isEmpty() || RequestNumber == ""))
//			ExcelFunctions.writeDataAtCell(InputDataFile,InputSheetname, iExcelRow, RequestNumbercolumn,RequestNumber);
	}
	
	@Test (priority=4)
	public void tearDown()
	{
		try {
			SNOWLog.info("Closing Browser - " + driver);
			if (driver != null)
				driver.close();

		} catch (Exception e) {
			SNOWLog.info(e.getMessage());
		} 
		
	}
	
	public static void main(String args[]) throws Exception
    {
    	try
    	{
    		//DOMConfigurator.configure("log4j.xml");
    		SNOWLog = LogFileCreator("C:/PCS RPA/DigitalInvoicing/LogFile/Log/RaiseSNOWInvoiceRequest.log", snowMassUploadFullAutomation.class.getSimpleName());
    		TestListenerAdapter tla = new TestListenerAdapter();
			TestNG testng = new TestNG();
			testng.setTestClasses(new Class[] { snowMassUploadFullAutomation.class });
			testng.addListener(tla);
			testng.run();
    	}
    	catch(Exception e)
    	{
    		throw e;
	    }
	}
	
}