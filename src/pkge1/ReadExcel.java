package pkge1;



import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import atu.testng.reports.ATUReports;

public class ReadExcel { 
	WebDriver driver;
	private static final String DELIM = "@";
	public static final String DRIVEREXCEL = "E:\\excel files\\Automationdriver.xls";
	public static final String RESULTEXCEL = "E:\\excel files\\Autoresult.xls";
	public static final String WEBEXP = "WebDriver Exception is thrown";
	public static final String EXP = "Unknown Exception is thrown";
	public static final String THREADEXP = "UnInteruptted Exception is thrown";
	public static final String NOELEEXP = "NoSuchElement Exception is thrown";
	public static final String IOEXP = "IO Exception is thrown";
	public static final String NIL = "";
	public static Workbook wb;
	public static Sheet sheet;
	WriteExcel we = new WriteExcel();
	public static String[][] tables;
	
	public ReadExcel(WebDriver webDriver) { 
		super();
		this.driver = webDriver; 

	}
	
	public void readExcelSheet(String destFile, String ShName) throws Exception {
		
		
		/*
		DesiredCapabilities capabilities = DesiredCapabilities.firefox();
		capabilities.setCapability(CapabilityType.PROXY, proxy);

		driver = new FirefoxDriver(capabilities);*/	
		
		
			TestEngine te = new TestEngine(driver);
			Workbook wb = Workbook.getWorkbook(new File(destFile));
			Sheet sheet = wb.getSheet(ShName);
			int rows = sheet.getRows();
			 String met = null;
			String prop = null;
			String data = null;
			for (int row = 0; row < rows; row++) {
				
				met = sheet.getCell(0, row).getContents();
				prop = sheet.getCell(1, row).getContents();
				data = sheet.getCell(2, row).getContents();
				if(!met.equalsIgnoreCase("TC")){
								
				Workbook wb1 = Workbook.getWorkbook(new File(RESULTEXCEL));
				Sheet sheet1 = wb1.getSheet(0); 
				int rows1 = sheet1.getRows();
				String chkExp1 = sheet1.getCell(2, rows1-1).getContents();
				if(chkExp1.equals(NIL))
				{
					driver.quit();
					ATUReports.add(met, prop, data, "", false);
					break; 
					
				}				 
				}
				
				WriteExcel we = new WriteExcel();
				switch (met) {
				
				case "launchApp":
					te.launchApp();
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "clickByid":
					te.clickByid(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;		
					
				case "enterByid":
					te.enterByid(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "clickDeny":
					te.clickDeny(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "clickApprove":
					te.clickApprove(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
					
				case "TC":
					String[] arrDT = data.split(DELIM);
					we.writeResult(RESULTEXCEL, arrDT[0], prop, arrDT[1]);
					ATUReports.add(met, prop, data, "", false);
					break;						
				
				case "moveByXpath" :
					te.moveByXpath(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "clickByXpath" :
					te.clickByXpath(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "clickByLink" :
					te.clickByLink(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
							
				case "clickByXPath" :
					te.clickByXPath(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
			
				case "enterByxPath" :
					te.enterByxPath(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
									
				case "verifyTextBoxes" :
					te.verifyTextBoxes(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "clickByCSS" :
					te.clickByCSS(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;		
							
				case "mouseOver" :
					te.mouseOver(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "mouseOverregister" :
					te.mouseOverregister(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
				
				case "clickByclassname" :
					te.clickByclassname(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "verifyText" : 
					te.verifyText(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "verifyTextS1" : 
					te.verifyTextS1(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "verifyTextipm" : 
					te.verifyTextipm(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
					
				case "movetoElement" :  
					te.movetoElement(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "movetoElementdropdown" : 
					te.movetoElementdropdown(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	 
					
				case "navigateto" : 
					te.navigateto(prop,data);  
					ATUReports.add(met, prop, data, "", false);
					break;	

				case "clickonPublish" : 
					te.clickonPublish(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;	

				case "recentAchievement" : 
					te.recentAchievement(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;

				case "scrollUp" : 
					te.scrollUp(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;


				case "achievementType" : 
					te.achievementType(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "tSleep" :
					te.tSleep(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Verifyearnedachievement" :
					te.Verifyearnedachievement(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "close" :
					te.close(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
					
				case "scrollDown" :
					te.scrollDown(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Delete" :
					te.Delete(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;

					
				case "logout" :
					te.logout(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Del" :
					te.Del(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "userresult" :
					te.userresult(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Password" :
					te.Password(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "achieveResult" :
					te.achieveResult(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "badgeLogin" :
					te.badgeLogin(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
					 
				case "badgeLogout" : 
					te.badgeLogout(prop,data); 
					ATUReports.add(met, prop, data, "", false);
					break;
								
				case "refresh" :
					te.refresh(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "submit":
					te.submit(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "scoreResult": 
					te.scoreResult(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "ResultAchievementType": 
					te.ResultAchievementType(prop,data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "ParameterDelete":  
					te.ParameterDelete(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Parameter":  
					te.Parameter(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Dropdown":   
					te.Dropdown(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "Dropdownbyxpath":   
					te.Dropdownbyxpath(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "navigateclassname":   
					te.navigateclassname(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "submitText":   
					te.submitText(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "verifyResultSucessMessage":   
					te.verifyResultSucessMessage(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "verifyResultSucessMessage1":    
					te.verifyResultSucessMessage1(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "verifyResultSucessMessage2":     
					te.verifyResultSucessMessage2(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					
				case "addComments":    
					te.addComments(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;
					 
				case "verifyAddcommentsSucessMessage":    
					te.verifyAddcommentsSucessMessage(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "navigatetoUrl":    
					te.navigatetoUrl(prop, data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "navigatetoprofile":     
					te.navigatetoprofile(prop, data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "verifyStudent":      
					te.verifyStudent(prop, data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "refreshcurrentpage":    
					te.refreshcurrentpage(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
						
				case "deActivatePlugins":     
					te.deActivatePlugins(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					 
				case "activatePlugins":     
					te.activatePlugins(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
					
				case "pageVerificationResult":      
					te.pageVerificationResult(prop, data); 
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "S1Result":        
					te.S1Result(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break; 
					
				case "S2Result":        
					te.S2Result(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "DenyResult":        
					te.DenyResult(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "ApproveResult":         
					te.ApproveResult(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	

				case "verifyS3result":       
					te.verifyS3result(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
				case "verifyAddcommentsResult":       
					te.verifyAddcommentsResult(prop, data);
					ATUReports.add(met, prop, data, "", false);
					break;	
					
					 
				default:
					ATUReports.add(met, prop, data, "", false);
					break;

				}
			}
	}
	public static Object[][] getTableArray() throws BiffException, IOException{
		int i=0;
		wb = Workbook.getWorkbook(new File(DRIVEREXCEL));           
		sheet = wb.getSheet(0);
		int rows = sheet.getRows();
		tables = new String[rows][3];
		System.out.println(rows);
		for (int row = 1; row < rows; row++) { 
			tables[i][0] = sheet.getCell(2, row).getContents();
			tables[i][1] = sheet.getCell(3, row).getContents();
			tables[i][2] = sheet.getCell(1, row).getContents();
			i++;
		}
		return tables;
	}
}
