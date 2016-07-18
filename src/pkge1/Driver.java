package pkge1; 
              

import java.awt.AWTException;
import java.io.File; 
import java.io.IOException;
import java.text.SimpleDateFormat; 
import java.util.Date;

import jxl.read.biff.BiffException;    
import jxl.write.WriteException;  

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;

import atu.testng.reports.ATUReports;
import atu.testng.reports.listeners.ATUReportsListener;
import atu.testng.reports.listeners.ConfigurationListener;
import atu.testng.reports.listeners.MethodListener;
import atu.testng.reports.logging.LogAs;
import atu.testng.reports.utils.Utils;
import atu.testng.selenium.reports.CaptureScreen;
import atu.testng.selenium.reports.CaptureScreen.ScreenshotOf;

@Listeners({ ATUReportsListener.class, ConfigurationListener.class, MethodListener.class })
                            
public class Driver {  
	 //Set Property for ATU Reporter Configuration
    {
      System.setProperty("atu.reporter.config", "C:\\Users\\NatarajanRamanathan\\NeonWorkspace\\BadgeOSKeyword\\src\\atu.properties");
    }
    public static WebDriver driver;
	public static final String DRIVEREXCEL = "E:\\excel files\\Automationdriver.xls";
	public static final String RESULTEXCEL = "E:\\excel files\\Autoresult.xls";   
	public static final String WEBEXP = "WebDriver Exception is thrown"; 
	public static final String EXP = "Unknown Exception is thrown";
	public static final String THREADEXP = "UnInteruptted Exception is thrown"; 
	public static final String NOELEEXP = "NoSuchElement Exception is thrown"; 
	public static final String IOEXP = "IO Exception is thrown";
	public static final String WRITEEXP = "Write Exception is thrown";
	public static final String BIFFEXP = "Unable to recognize OLE stream";      
	static final String AWTEXP = "AWT exception is thrown"; 
	public static final String NIL = "";  
	public Object DataTable[][];
	//public WebDriver driver;	 
	private void setAuthorInfoForReports() {
        ATUReports.setAuthorInfo("Automation Tester", Utils.getCurrentTime(),"1.0");
	}
	@BeforeClass
	public void init(){
		driver = new FirefoxDriver();
		//System.setProperty("webdriver.chrome.driver", "C://Selenium//chromedriver.exe");
		//WebDriver driver = new ChromeDriver();
		ATUReports.setWebDriver(driver);
	 	ATUReports.indexPageDescription = "<br> Please change this <br/> <b>Can include Full set of HTML Tags</b>";
	}
	@AfterClass
	public void close(){
		driver.quit();
	}
	@Test(dataProvider="Driver")
	public void test(String Exe,String ShName, String description) {   
		setAuthorInfoForReports();
		WriteExcel we = new WriteExcel();
				try {                                                     
					                                                                   
			// driver = new FirefoxDriver();              
			// dr.moveTheFile();           
			if (Exe.equalsIgnoreCase("Yes")) {  
				//dr.openBrowser();  
				try {
					ReadExcel re = new ReadExcel(driver);  
					re.readExcelSheet(DRIVEREXCEL, ShName);             
					 
				}      
				            
				     
		 		catch (Exception e) {  
					if(e instanceof WriteException){ 
						we.writeResult(RESULTEXCEL, WRITEEXP, NIL, NIL); 
					}else if (e instanceof BiffException) { 
						we.writeResult(RESULTEXCEL, BIFFEXP, NIL, NIL);	 
						
					}else if (e instanceof IOException) {   
						  
						we.writeResult(RESULTEXCEL, IOEXP, NIL, NIL);          
						  
					}else if (e instanceof AWTException) {
						we.writeResult(RESULTEXCEL, AWTEXP, NIL, NIL);
						 
					}else if (e instanceof org.openqa.selenium.NoSuchElementException) { 
						we.writeResult(RESULTEXCEL, NOELEEXP, NIL, NIL);
					}else if (e instanceof InterruptedException) {
						we.writeResult(RESULTEXCEL, THREADEXP, NIL, NIL);
					}else if (e instanceof WebDriverException ) { 
						we.writeResult(RESULTEXCEL, WEBEXP, NIL, NIL);	   
					}else { 
						we.writeResult(RESULTEXCEL, EXP, NIL, NIL);	            
					}  
				ATUReports.add("", LogAs.FAILED	, new CaptureScreen(ScreenshotOf.BROWSER_PAGE));
				}     
				
			}   
		}	                      
			catch (Exception e) {       
       
			e.printStackTrace();         
			ATUReports.add("", LogAs.FAILED	, new CaptureScreen(ScreenshotOf.BROWSER_PAGE));
		} 
				
	}  
		//////system.out.println(rows);
	@DataProvider(name="Driver")
	public Object[][] driver() throws BiffException, IOException
	{
		DataTable = ReadExcel.getTableArray();
		return DataTable;
	}
	
			 
			
//*******************************************************************************************************
	
	
	
	
	public void moveTheFile() throws WriteException, BiffException, IOException {
			File ExeDir = new File("K:\\congruent\\Execution");
			if (!ExeDir.exists()) {
				ExeDir.mkdir();
			} 
			File SnapDir = new File("K:\\congruent\\Execution\\Snapshots");
			if (!SnapDir.exists()) {
				SnapDir.mkdir();   
			} 
 
			File srcTD = new File("K:\\congruent\\Master\\Driver.xls");
			File srcTR = new File("K:\\congruent\\Master\\TestResult.xls");
			File dstTD = new File("K:\\congruent\\Execution\\Driver.xls");
			if (!dstTD.exists()) { 
					FileUtils.copyFileToDirectory(srcTD, ExeDir);  
			}
			File dstTR = new File("K:\\congruent\\Execution\\TestResult.xls");
			if (!dstTR.exists()) {
			 
					FileUtils.copyFileToDirectory(srcTR, ExeDir);   
			}
	    
	}    

	// Archiving the Files and Folders to Backup folder with Date and Timestamp
	public void archiveFile() throws IOException {
			File srcDir = new File("K:\\congruent\\Execution");
			SimpleDateFormat dateFormat = new SimpleDateFormat(   
					"dd-MM-yyyy-HH-mm-ss");        
			Date date = new Date();  
			String dt = dateFormat.format(date);   
			// ////system.out.println(dt);
			File dtbk = new File("K:\\congruent\\Archived\\back up (" + dt + ")");
			FileUtils.copyDirectory(srcDir, dtbk);
			FileUtils.deleteDirectory(srcDir);	
	}  

	/*
	public void screenshot(String TS, String TestStepID) throws IOException { 
		
		//WebDriver driver = new FirefoxDriver(); 
	   	File scrFile2 = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);  
			
			//FileUtils.copyFile(scrFile2, new File("D:\\Autofiles\\Screenshot\\Screenshot" + TS + "_"+ TestStepID + ".png"));
			FileUtils.copyFile(scrFile2, new File("D:\\Autofiles\\Screenshot\\Screenshot" + ".png"));
			//FileUtils.copyFile(scrFile2, new File("D:\\Autofiles\\Screenshot\\Screenshot.png"));
	}
*/
}