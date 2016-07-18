package pkge1;

 
import java.awt.AWTException;
import java.io.File;
import java.io.IOException;
import java.util.List; 
import java.util.concurrent.TimeUnit;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.commons.io.FileUtils;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;


public class TestEngine  {   
	public WebDriver driver;
	private static final String DELIM = "@"; 
	private static final String URL = "http://dev.asahitechnologies.com/badgeos/"; 
	public static final String RESULTEXCEL = "E:\\excel files\\Autoresult.xls";
	public static final String DRIVEREXCEL = "E:\\excel files\\Automationdriver.xls";
	//public static final String INPUTEXCEL = "D:\\karthee\\InputFiles\\Data1.xls";
	public static final String INPUTEXCEL = "D:\\BadgeOsAutomationFiles\\Testdata.xls";
	public static final String WEBEXP = "WebDriver Exception is thrown";
	public static final String EXP = "Unknown Exception is thrown";
	public static final String THREADEXP = "UnInteruptted Exception is thrown";
	public static final String NOELEEXP = "NoSuchElement Exception is thrown";
	public static int fl = 0;
	public static int fp = 0;
	public static final String NIL = "";
	WriteExcel writeExcel = new WriteExcel();
	Driver objDriver = new Driver();

	// Constructor to initialize the driver
	public TestEngine(WebDriver webDriver) { 
		super();
		this.driver = webDriver; 

	}

	public void launchApp() throws InterruptedException {
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
			driver.get(URL);
			driver.manage().window().maximize();
			  
	
	}

	public void clickByid(String pr, String dt) throws InterruptedException {
		System.out.println("ok");
			driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
			Thread.sleep(1000);
			if (driver.findElement(By.id(pr)).isEnabled()) {
				//system.out.println("button enabled");
				driver.findElement(By.id(pr)).click();
			}
	}
	
	public void enterByid(String pr, String dt) throws InterruptedException {
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.findElement(By.id(pr)).click();
		driver.findElement(By.id(pr)).clear();
		Thread.sleep(500);
		driver.findElement(By.id(pr)).sendKeys(dt);
		//driver.findElement(By.id(pr)).sendKeys(Keys.TAB);
	}
	//verify a page name
	
	public void enterByid1(String pr, String dt) throws InterruptedException {
		
		
		driver.findElement(By.id(pr)).clear();
		driver.findElement(By.id(pr)).sendKeys(dt);
	}

	public void enternewItemCode(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		Thread.sleep(5000);
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		//Arrdt = dt.split("%");
		driver.findElement(By.cssSelector(Arrpr[0])).click();
		driver.findElement(By.cssSelector(Arrpr[1])).sendKeys(dt);

	}
	
	public void clicknewComponentType(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		//Arrdt = dt.split("%");
		
		WebDriverWait ww = new WebDriverWait(driver, 10);
		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(Arrpr[0])));
		
	
		driver.findElement(By.cssSelector(Arrpr[0])).click();
	
		//System.out.println(driver.findElement(By.cssSelector(Arrpr[3])).isDisplayed());
		List<WebElement> we = driver.findElements(By.cssSelector(Arrpr[1]));
		for(WebElement wee: we)
		{

			List<WebElement> weul = wee.findElements(By.tagName("ul"));
			
			for(int i=0;i<weul.size();i++)
			{
				List<WebElement> weli = wee.findElements(By.tagName("li"));
				for(int j=0;j<weli.size();j++)
				{
					String str = weli.get(j).getText();
					if(str.equalsIgnoreCase(dt))
					{
						weli.get(j).click();
						break;
					}
				
				}
			
			}
		
		}
	}

	public void clicknewCountry(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr;
		Arrpr = pr.split("%");
		
		Thread.sleep(5000);

		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Country"))
		{
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			WebElement weclickcntry = driver.findElement(By.cssSelector(Arrpr[3]));
			
			List<WebElement> weclickcntry2 = weclickcntry.findElements(By.cssSelector("div.select-list-container"));
			
			
			
			
			for(WebElement wee: weclickcntry2)
			{
				System.out.println("am here1");
				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							System.out.println(j);
							weli.get(j).click();
							break;
							
							
						}
					}
				}
			}
			
		}
	}
	
	

	public void clicknewProductName(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		
		System.out.println("am hereeeeeeeeeee");
		String[] ArrPolpr;
		
		ArrPolpr = pr.split("%");

		Thread.sleep(5000);
		
		System.out.println("am here");
		
		WebElement we = driver.findElement(By.xpath(ArrPolpr[0]));
		
		WebElement welab = we.findElement(By.xpath(ArrPolpr[1]));
		
		String str = welab.getText();
		
		System.out.println(str);
		
		System.out.println((str.equals("Product Name")));
		
		if(str.equals("Product Name"))
		{
			WebElement weclickcntry1 = driver.findElement(By.xpath(ArrPolpr[2]));
			
			weclickcntry1.click();
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(ArrPolpr[3]));
			
			for(WebElement wee: weclickcntry)
			{
				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					}
				}	
			}
			
		}
	}
	
	public void clicknewProductName1(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] ArrPolpr;
		ArrPolpr = pr.split("%");
		//ArrPolpr
		Thread.sleep(5000);
	
		WebElement we = driver.findElement(By.xpath(ArrPolpr[0]));
		
		WebElement welab = we.findElement(By.xpath(ArrPolpr[1]));
		
		String str = welab.getText();
		
		System.out.println(str);
		
		System.out.println((str.equals("Product Name")));
		
		if(str.equals("Product Name"))
		{
			WebElement weclickcntry1 = driver.findElement(By.xpath(ArrPolpr[2]));
			
			weclickcntry1.click();
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(ArrPolpr[3]));
			for(WebElement wee: weclickcntry)
			{
				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						
						System.out.println("1");
						System.out.println(str1);
						
						System.out.println(dt);
						
						System.out.println("2");
						if(str1.equalsIgnoreCase(dt))
						{
							

							weli.get(j).click();
							
							
							System.out.println(driver.findElement(By.xpath(ArrPolpr[4])).isDisplayed());
							WebElement we1 = driver.findElement(By.xpath(ArrPolpr[4]));
							
							WebElement welab1 = we1.findElement(By.xpath(ArrPolpr[5]));
							
							String strapi = welab1.getText();
							
							System.out.println(strapi);
							
							System.out.println(strapi.equalsIgnoreCase("API"));
							
							
							WebElement weapele = driver.findElement(By.xpath(ArrPolpr[10]));
							
							
							WebElement weapele1 = weapele.findElement(By.xpath(ArrPolpr[11]));
							
							String strapis = weapele1.getText();
						
							
							
							if(strapis.equals("API"))
							{
								
								System.out.println(driver.findElement(By.xpath(ArrPolpr[12])).isDisplayed());
								WebElement weclickcntry11 = driver.findElement(By.xpath(ArrPolpr[12]));
								weclickcntry11.click();
								
								String strapitxts = driver.findElement(By.xpath(ArrPolpr[12])).getText();
								
								List<WebElement> weclickcntry13 = driver.findElements(By.cssSelector(ArrPolpr[13]));
								for(WebElement wee13: weclickcntry13)
									
								{
									
									List<WebElement> wee13ele = wee13.findElements(By.tagName("ul"));
									
									for(int ii=0;ii<weul.size();ii++)
										
									{
										List<WebElement> wee13elelis = wee13.findElements(By.tagName("li"));
										
										for(int jj=0;jj<weli.size();jj++)
											
										{
											
											
											String strapi1 = wee13elelis.get(jj).getText();
											
											System.out.println(strapi1);
							
											if(strapi.equalsIgnoreCase("API"))
												
											{
												System.out.println("Pass");
												//writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the API is displaying or not for the Product - " +ArrPoldt[0],strapi,strapi+ " - APIs are displaying for the Product Name - " +ArrPoldt[0],strapi+ " - APIs are displaying for the Product Name - " +str1, "Pass");
												writeExcel.writeResult(RESULTEXCEL, ArrPolpr[8],ArrPolpr[9],"Verify the API is displaying or not for the Product - " +dt,strapitxts,strapitxts+ " - APIs are displaying for the Product Name - " +dt,strapitxts+ " - APIs are displaying for the Product Name - " +str1, "Pass");
												break;
											}
											
										//	break;
											else
											{
												System.out.println("Fail");
												writeExcel.writeResult(RESULTEXCEL, ArrPolpr[8],ArrPolpr[9],"Verify the API is displaying or not for the Product - " +dt,strapi,strapi+ " - APIs are not displaying for the Product Name - " +dt,strapi+ " - APIs are not displaying for the Product Name - " +str1, "Pass");
											}
										}
									}break;
								}
							}break;
							
						}//break;
					
					}
				
				}break;
			
			}
			
		}
	}
	
	
	public void verifyRule(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] ArrPolpr, ArrPoldt;
		ArrPolpr = pr.split("%");
		ArrPoldt = dt.split("%");
		
		Thread.sleep(5000);
		
		WebElement we = driver.findElement(By.xpath(ArrPolpr[4]));
		
		
		WebElement welab = we.findElement(By.xpath(ArrPolpr[5]));
		
		String str = welab.getText();
		
		System.out.println(str);
		
		System.out.println((str.equals("Product Name")));
		
		if(str.equals("Product Name"))
		{
			//WebElement weclickcntry1 = driver.findElement(By.xpath(ArrPolpr[2]));
			//weclickcntry1.click();
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(ArrPolpr[6]));
			for(WebElement wee: weclickcntry)
			{
				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							//weli.get(j).click();
							
							WebElement weapele = driver.findElement(By.xpath(ArrPolpr[10]));
							
							
							WebElement weapele1 = weapele.findElement(By.xpath(ArrPolpr[11]));
							
							String strapi = weapele1.getText();
							
							if(str.equals("API"))
							{
								
								
								WebElement weclickcntry1 = driver.findElement(By.xpath(ArrPolpr[12]));
								weclickcntry1.click();
										
								
								
								List<WebElement> weclickcntry13 = driver.findElements(By.cssSelector(ArrPolpr[13]));
								for(WebElement wee13: weclickcntry13)
								{

									List<WebElement> wee13ele = wee13.findElements(By.tagName("ul"));
									
									for(int ii=0;ii<weul.size();ii++)
									{
										List<WebElement> wee13elelis = wee13.findElements(By.tagName("li"));
										for(int jj=0;jj<weli.size();jj++)
										{
											String strapi1 = wee13elelis.get(jj).getText();
							
											if(strapi.equalsIgnoreCase("API"))
											{
												System.out.println("Pass");
												//writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the API is displaying or not for the Product - " +ArrPoldt[0],strapi,strapi+ " - APIs are displaying for the Product Name - " +ArrPoldt[0],strapi+ " - APIs are displaying for the Product Name - " +str1, "Pass");
												writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the API is displaying or not for the Product - " +dt,strapi,strapi+ " - APIs are displaying for the Product Name - " +dt,strapi+ " - APIs are displaying for the Product Name - " +str1, "Pass");
								
											}
											break;
										}
									}
								}
							}
							
							
							
							
						}
					
					}
				
				}
			
			}
			
		}
		
	}
	
	
	
	
	
	public void clicknewAPI(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] Arrpr;
		Arrpr = pr.split("%");
		Thread.sleep(10000);
		
		System.out.println("am here");
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		System.out.println("am here");
		
		String str = welab.getText();
		
		if(str.equals("API"))
		{
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{
				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						System.out.println("am here");
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					}
				}
			}
		}
	}
	
	public void clicknewStrenght(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		
		//Thread.sleep(5000);
//		WebDriverWait ww = new WebDriverWait(driver, 10);
//		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(Arrpr[0])));
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Strength"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	public void clicknewPacksize(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		
		//Thread.sleep(5000);
//		WebDriverWait ww = new WebDriverWait(driver, 10);
//		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(Arrpr[0])));
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Pack Size"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	
	public void clicknewPharmaceuticalForm(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		
		//Thread.sleep(5000);
//		WebDriverWait ww = new WebDriverWait(driver, 10);
//		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(Arrpr[0])));
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Pharmaceutical Form"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	
	

	public void clicknewArtworkRequestType(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		
		//Thread.sleep(5000);
//		WebDriverWait ww = new WebDriverWait(driver, 10);
//		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(Arrpr[0])));
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Artwork Request Type"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	
	

	public void clicknewPharmacode(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		Arrdt = dt.split("%");
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		System.out.println("12121");
		
		if(str.equals("Pharmacode"))
		{
			System.out.println(driver.findElement(By.xpath(Arrpr[2])).isDisplayed());
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			//driver.findElement(By.xpath(Arrpr[3])).click();
			
			System.out.println(driver.findElement(By.xpath(Arrpr[2])).getText());
			
				Actions act = new Actions(driver);
				Select se = new Select(weclickcntry1);
				se.selectByVisibleText(Arrdt[0]);
				act.click().build().perform();
				
				act.sendKeys(Keys.ENTER).build().perform();
				act.sendKeys(Keys.TAB).build().perform();
				System.out.println(driver.findElement(By.xpath(Arrpr[4])).isDisplayed());
				//driver.findElement(By.cssSelector(Arrpr[4])).click();
				act.sendKeys(Arrdt[1]).build().perform();
				act.sendKeys(Keys.TAB).build().perform();
			
		}
	}
	
	
	
	
	public void clicknewBarcode(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		Arrdt = dt.split("%");
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		WebElement we1 = we.findElement(By.xpath(Arrpr[1]));
		WebElement we2 = we1.findElement(By.xpath(Arrpr[2]));
		we2.click();
		List<WebElement> we3 = we1.findElements(By.xpath(Arrpr[2]));
		
		for(int i=0;i<we3.size();i++)
		{
			Select sel = new Select(we3.get(i));
			sel.selectByVisibleText(Arrdt[0]);
			Actions act = new Actions(driver);
			act.sendKeys(Keys.ENTER).build().perform();
			
			
			act.sendKeys(Keys.TAB).build().perform();
			
		//	we3.get(i).findElement(By.xpath("div[1]/div[2]/div[4]/div/div/div[2]"));
			
			act.sendKeys(Arrdt[1]).build().perform();
			act.sendKeys(Keys.TAB).build().perform();
		}
		
	}
	
	
	
	public void clicknewPackagingSite(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		Thread.sleep(10000);
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		String str = welab.getText();
		System.out.println(str);
		if(str.equals("Packaging Site"))
		{
			System.out.println(driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div[1]/div/div")).isDisplayed());
			driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div[1]/div/div")).click();
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{
				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	
	public void clicknewBrailleRequired(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		Thread.sleep(5000);
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Braille Required"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	
	
	public void enternewBrailleTranslation(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Braille Translation"))
		{
			System.out.println("am here");
			
			System.out.println(driver.findElement(By.cssSelector(Arrpr[2])).isDisplayed());
			
			WebElement weclickcntry1 = driver.findElement(By.cssSelector(Arrpr[2]));
			weclickcntry1.click();
					
			driver.findElement(By.cssSelector(Arrpr[2])).sendKeys(dt);;
			
			
		}
	}
	
	
	public void clicknewPerformation(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Performation"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	
	
	public void clicknewGraphicAgency(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		
		
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		
		
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		
		String str = welab.getText();
		
		if(str.equals("Global Graphic Agency"))
		{
			
			
			WebElement weclickcntry1 = driver.findElement(By.xpath(Arrpr[2]));
			weclickcntry1.click();
					
			
			
			List<WebElement> weclickcntry = driver.findElements(By.cssSelector(Arrpr[3]));
			for(WebElement wee: weclickcntry)
			{

				List<WebElement> weul = wee.findElements(By.tagName("ul"));
				
				for(int i=0;i<weul.size();i++)
				{
					List<WebElement> weli = wee.findElements(By.tagName("li"));
					for(int j=0;j<weli.size();j++)
					{
						String str1 = weli.get(j).getText();
						if(str1.equalsIgnoreCase(dt))
						{
							weli.get(j).click();
							break;
						}
					
					}
				
				}
			
			}
			
		}
	}
	

	public void clicknewBrowseButton(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		WebElement we = driver.findElement(By.cssSelector(Arrpr[0]));
		
		we.sendKeys("D:\\Users\\kartheeswaran.r\\Desktop\\pdf\\C0007.01.pdf");
		
		Actions act = new Actions(driver);
		act.sendKeys(Keys.TAB).build().perform();
		Thread.sleep(10000);
		
		
		
		
		
		
		
		driver.findElement(By.cssSelector(Arrpr[1])).click();
		
		
		
	}
	
	
	
	public void enternewInstructions(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr, Arrdt;
		Arrpr = pr.split("%");
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		WebElement welab = we.findElement(By.xpath(Arrpr[1]));
		String str = welab.getText();
		if(str.equals("Instructions"))
		{
			WebElement weclickcntry1 = driver.findElement(By.cssSelector(Arrpr[2]));
			weclickcntry1.click();
			driver.findElement(By.cssSelector(Arrpr[2])).sendKeys(dt);;
			
		}
	}
	
	
	
	
	public void dragRouting1(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] Arrpr,Arrdt;
		Arrpr = pr.split("%");
		Arrdt = dt.split("%");
		Thread.sleep(5000);
//		driver.findElement(By.xpath(Arrpr[0])).sendKeys(Arrdt[0]);
//		driver.findElement(By.cssSelector(Arrpr[1])).click();
		Thread.sleep(10000);
		driver.switchTo().frame(0);
		driver.switchTo().frame(driver.findElement(By.cssSelector(Arrpr[2])));
		driver.findElement(By.xpath(Arrpr[3])).click();
		List<WebElement> wee = driver.findElements(By.cssSelector(Arrpr[4]));
		for(WebElement weee:wee)
		{
			String str = weee.getText();
			if(str.equals(Arrdt[1]))
			{
					WebElement wee1 = driver.findElement(By.cssSelector(Arrpr[5]));
					Actions act = new Actions(driver);
					act.moveToElement(weee).build().perform();
					act.clickAndHold(weee).build().perform();
					act.release(wee1).build().perform();
					break;
			}
		
		
		// ********************** " Save as Template"
	//	driver.findElement(By.cssSelector(Arrpr[6])).click();
	//	driver.findElement(By.cssSelector(Arrpr[7])).sendKeys(Arrdt[2]);
	//	driver.findElement(By.cssSelector(Arrpr[8])).click();
			
			System.out.println("dragged");
			Thread.sleep(5000);
		driver.findElement(By.cssSelector(Arrpr[9])).click();
		Thread.sleep(5000);
		break;
		}
			//break;
		
	}
	
	
	public void dragRouting(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr,Arrdt;
		Arrpr = pr.split("%");
		Arrdt = dt.split("%");
		/*//driver.switchTo().frame(0);
		System.out.println(driver.findElement(By.xpath("//*[@ng-model='filter.search']")).isDisplayed());
		Thread.sleep(5000);
		driver.findElement(By.xpath("//*[@ng-model='filter.search']")).sendKeys(" Test11");
		Thread.sleep(5000);
		driver.findElement(By.xpath("//*[@href='Request?Type=Task&taskID=135&Frame=true']")).click();
		
		driver.switchTo().frame(0);
		System.out.println(driver.findElement(By.xpath("//*[@value='Submit']")).isDisplayed());
		driver.findElement(By.xpath("//*[@value='Submit']")).click();
		driver.switchTo().defaultContent();*/
		Thread.sleep(5000);
		
		
		driver.findElement(By.xpath(Arrpr[0])).sendKeys(Arrdt[0]);
		
		
		driver.findElement(By.cssSelector(Arrpr[1])).click();
		
		Thread.sleep(10000);
		
		driver.switchTo().frame(0);
		
		driver.switchTo().frame(driver.findElement(By.cssSelector(Arrpr[2])));
		
	//	System.out.println(driver.findElement(By.xpath("/html/body/form/div/div[3]/ul[1]/li[1]/a")).isDisplayed());
		
		//System.out.println("1");
		
		driver.findElement(By.xpath(Arrpr[3])).click();
		
		//System.out.println(driver.findElement(By.cssSelector("li.ui-draggable-handle[ng-repeat='r in Data.RoutingTemplate']")).isDisplayed());
		
		List<WebElement> wee = driver.findElements(By.cssSelector(Arrpr[4]));
		
		for(WebElement weee:wee)
		{
			String str = weee.getText();
		
				if(weee.getText().equals(Arrdt[1]))
				{
			
					WebElement wee1 = driver.findElement(By.cssSelector(Arrpr[5]));
					Actions act = new Actions(driver);
					act.moveToElement(weee).build().perform();
					act.clickAndHold(weee).build().perform();
					act.release(wee1).build().perform();
					break;
				}
		}
		
		
		
		
		driver.findElement(By.cssSelector(Arrpr[6])).click();
		
		driver.findElement(By.cssSelector(Arrpr[7])).sendKeys(Arrdt[2]);
		
				
		
		
		driver.findElement(By.cssSelector(Arrpr[8])).click();
		
		
		
		driver.findElement(By.cssSelector(Arrpr[9])).click();
		
		
		
		
		
		
		
		
		
		
	/*	
		
		List<WebElement> iframes = driver.findElements(By.tagName("iframe"));
		for (WebElement iframe : iframes) {
			System.out.println("1");
			
			System.out.println(iframe.findElement(By.xpath("/html/body/form/div/div[2]/div[1]/svg/rect[2]")).isDisplayed());
		}
	       
		System.out.println("1");
		
		Thread.sleep(5000);
		
			
	System.out.println(driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div/a[2]")).isDisplayed());
		System.out.println(driver.findElement(By.cssSelector("html body form div.ng-scope div.footer.text-center")).isDisplayed());
		WebElement we = driver.findElement(By.cssSelector("div.footer"));
		List<WebElement> weul = we.findElements(By.tagName("ul"));
		for(int i=0;i<weul.size();i++)
		{
			
			List<WebElement> weli = we.findElements(By.tagName("li"));
			for(int j=0;j<weli.size();j++)
			{
				System.out.println(weul.get(j).getText());
			}
		}
		
		
		System.out.println(driver.findElement(By.xpath("/html/body/form/div/div[3]/ul[1]/li[1]/a")).isDisplayed());
		
		
		driver.findElement(By.xpath("//*[@ng-click='toggleRoutingTemplate = !!!toggleRoutingTemplate']")).click();
		*/
	}
	
	
	
	
	

	
	public void createMasterDt(String pr, String dt)throws Exception {
		
		
		String[] Arrpr;
		Arrpr = pr.split("%");
		//System.out.println("am here");
	
		//System.out.println(dt.equalsIgnoreCase("MasterData"));
		
		WebElement we = driver.findElement(By.xpath(Arrpr[4]));
		Actions act = new Actions(driver);
		act.moveToElement(we).build().perform();
		WebElement we1 = driver.findElement(By.xpath(Arrpr[5]));
		act.moveToElement(we1).build().perform();
		WebElement we2 = driver.findElement(By.xpath(Arrpr[6]));
		act.click(we2).build().perform();
		
		
		
		
		if (dt.equalsIgnoreCase("MasterData")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("MasterData");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data, alrd, isuserprf="";
			for (int row = 3; row < rows; row++) 
			{
				data = sheet.getCell(0, row).getContents();
				
				alrd = sheet.getCell(1, row).getContents();
				
				isuserprf = sheet.getCell(3, row).getContents();
				
				if(alrd.equalsIgnoreCase("Yes"))
				{
					
					driver.findElement(By.linkText(Arrpr[0])).click();
					
					Thread.sleep(5000);
					System.out.println(driver.findElement(By.cssSelector(Arrpr[1])).isDisplayed());
					
					driver.findElement(By.cssSelector(Arrpr[1])).sendKeys(data);
					if(isuserprf.equalsIgnoreCase("Yes"))
					{
						driver.findElement(By.xpath(Arrpr[2])).click();
					}
					
					driver.findElement(By.xpath(Arrpr[3])).click();
					Thread.sleep(5000);
				}
			}
			
		}
		
		
		
	}
	
	
	
	
public void enterByxPath(String pr, String dt) throws InterruptedException {
		
	//driver.switchTo().frame(0);
		//System.out.println(driver.findElement(By.xpath(pr)).isDisplayed());
	
	WebDriverWait ww = new WebDriverWait(driver, 10);
	ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
		driver.findElement(By.xpath(pr)).clear();
		driver.findElement(By.xpath(pr)).sendKeys(dt);
		
		
	}
	
	public void checkByXPath(String pr, String dt) throws InterruptedException {
		
		
		driver.findElement(By.xpath(pr)).click();
		
		
	}
		
	public void clickByXPath(String pr, String dt) throws InterruptedException, IOException {
		
		//Thread.sleep(5000);
		
		WebDriverWait ww = new WebDriverWait(driver, 10);
		ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
		
		driver.findElement(By.xpath(pr)).click();
		 
		//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		
	}
	
	public void clcikRoles(String pr, String dt) throws InterruptedException, BiffException, IOException {
	
		String[] Arrpr= pr.split("%");
		
		Actions act = new Actions(driver);
		
		//WebElement we3 = driver.findElement(By.xpath(Arrpr[0]));
		
		System.out.println("am here");
		
		System.out.println(Arrpr[0]);
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		System.out.println(Arrpr[1]);
		WebElement we1 = driver.findElement(By.xpath(Arrpr[1]));
		
		WebElement we2 = driver.findElement(By.xpath(Arrpr[2]));
		System.out.println(Arrpr[2]);
		WebElement we3 = driver.findElement(By.xpath(Arrpr[3]));
		
		
		
		
		
		act.moveToElement(we).build().perform();
		
		act.moveToElement(we1).build().perform();
		
		act.moveToElement(we2).build().perform();
		
		//act.moveToElement(we3).build().perform();
	//	Thread.sleep(5000);
		driver.findElement(By.linkText(Arrpr[4])).click();
		
		
		if (dt.equalsIgnoreCase("Roles")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("Roles");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data;
			for (int row = 0; row < rows; row++) 
			{
				data = sheet.getCell(0, row).getContents();
				
				driver.findElement(By.linkText(Arrpr[5])).click();
				
				driver.findElement(By.cssSelector(Arrpr[6])).sendKeys(data);
				
				//**************driver.findElement(By.cssSelector(Arrpr[7])).click();
				
				
				
			}
			
		}
}
	
	
	public void verifyRoles(String pr, String dt) throws InterruptedException, BiffException, IOException {
		
		String[] Arrpr= pr.split("%");
		
		Actions act = new Actions(driver);
		
		//WebElement we3 = driver.findElement(By.xpath(Arrpr[0]));
		
		System.out.println("am here");
		
		System.out.println(Arrpr[0]);
		WebElement we = driver.findElement(By.xpath(Arrpr[0]));
		System.out.println(Arrpr[1]);
		WebElement we1 = driver.findElement(By.xpath(Arrpr[1]));
		
		WebElement we2 = driver.findElement(By.xpath(Arrpr[2]));
		System.out.println(Arrpr[2]);
		WebElement we3 = driver.findElement(By.xpath(Arrpr[3]));
		
		
		
		
		
		act.moveToElement(we).build().perform();
		
		act.moveToElement(we1).build().perform();
		
		act.moveToElement(we2).build().perform();
		driver.findElement(By.linkText(Arrpr[4])).click();
		
		
		if (dt.equalsIgnoreCase("Roles")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("Roles");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data;
			for (int row = 0; row < rows; row++) 
			{
				data = sheet.getCell(0, row).getContents();
				
				List<WebElement> wetab = driver.findElements(By.className(Arrpr[8]));
				for(int i=0;i<wetab.size();i++)
				{
					String str = wetab.get(i).getText();
					System.out.println(str);
				
					
				}
			}
		}
	}

	
	public void dropdownSelect(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr;
		Arrpr = pr.split("%");
		//Thread.sleep(5000);
		Thread.sleep(2000);
		String mngdtnd, mngdt, mstrdt, mngdts = "";
		
		
		WebElement weet = driver.findElement(By.xpath(Arrpr[0]));
		Actions act1 = new Actions(driver);
		act1.moveToElement(weet).build().perform();
		driver.findElement(By.xpath(Arrpr[1])).click();
		
		
		
		if (dt.equalsIgnoreCase("ManagedData")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("ManagedData");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data;
			for (int row = 3; row < rows; row++) 
			{
				
				mstrdt = sheet.getCell(0, row).getContents();
				//System.out.println(mstrdt);
				mngdtnd = sheet.getCell(2, row).getContents();
				
				if(mngdtnd.equalsIgnoreCase("Yes"))
				{
					WebElement we = driver.findElement(By.cssSelector(Arrpr[2]));
					we.click();
					System.out.println(mstrdt);
					Select se = new Select(we);
					se.selectByVisibleText(mstrdt);
					act1.click().build().perform();
					act1.sendKeys(Keys.TAB).build().perform();
					
					
					Sheet sheet1 = wb.getSheet(mstrdt);
					int rows1 = sheet1.getRows();
					for (int row1 = 0; row1 < rows1; row1++) 
					{
						System.out.println(row1);
							mngdt = sheet1.getCell(0, row1).getContents();
								
							System.out.println(mngdt);
								Thread.sleep(2000);
								System.out.println("dai");
								System.out.println(driver.findElement(By.cssSelector(Arrpr[3])).isDisplayed());
								driver.findElement(By.cssSelector(Arrpr[3])).click();
								Thread.sleep(2000);
								driver.findElement(By.xpath(Arrpr[4])).sendKeys(mngdt);		
								driver.findElement(By.xpath(Arrpr[6])).click();
								
					}
					
				}
				
			}
		}

		
	}
	
	/*
	public void verifyMangedData(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] arrPolpr;
		arrPolpr = pr.split("%");
		
		//Thread.sleep(5000);
		String mngdtnd, mngdt, mstrdt, mngdts = "";
		
		WebElement weet = driver.findElement(By.xpath("//*[@title='Admin']"));
		Thread.sleep(2000);
		Actions act1 = new Actions(driver);
		act1.moveToElement(weet).build().perform();
		driver.findElement(By.xpath("//*[@title='Manage Data']")).click();
		
		
		
		if (dt.equalsIgnoreCase("ManagedData")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("ManagedData");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data;
			for (int row = 3; row < rows; row++) 
			{
				
				mstrdt = sheet.getCell(0, row).getContents();
				System.out.println(mstrdt);
				mngdtnd = sheet.getCell(2, row).getContents();
				
				if(mngdtnd.equalsIgnoreCase("Yes"))
				{
					WebElement we = driver.findElement(By.cssSelector("select.form-control[ng-model='ManagedDataType']"));
					we.click();
					Select se = new Select(we);
					se.selectByVisibleText(mstrdt);
					act1.click().build().perform();
					act1.sendKeys(Keys.TAB).build().perform();
					
					Sheet sheet1 = wb.getSheet(mstrdt);
					int rows1 = sheet1.getRows();
					for (int row1 = 0; row1 < rows1; row1++) 
					{
								
								mngdt = sheet1.getCell(0, row1).getContents();
								
							
								
								Thread.sleep(5000);
								
								System.out.println("am here");
								
								System.out.println(driver.findElement(By.cssSelector(arrPolpr[5])).isDisplayed());
								
								WebElement weall = driver.findElement(By.cssSelector(arrPolpr[5]));
								
								//System.out.println("am here");
								List<WebElement> weallele = weall.findElements(By.tagName("tr"));
								int len = weallele.size();
								int flag=0;
								for(int i = 0;i<len;i++)
								{	
									
									
									WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/form/div/div[2]/div[2]/table/tbody/tr['" + i + "'+1]/td[2]"));
									
									String strtxt = we5.getText();
									System.out.println(strtxt);
										System.out.println(strtxt.equalsIgnoreCase(mngdt));
										
										
										if(strtxt.equals(mngdt))
										{
											flag++;
											
											System.out.println("Pass");
											writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Managed Data",mngdt,mngdt+ " - Managed Data is available for the Master Data - '"+mstrdt+"'",strtxt+ " - Managed Data is available for the Master Data - '"+mstrdt+"'", "Pass");
											
											
											((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
											
											//((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
											
											objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
											break;
										}
								
										
									}
										System.out.println(flag);
										if(flag==0)
										{
											System.out.println(flag);
											System.out.println("fail");
											writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Managed Data",mngdt,mngdt+ " - Managed Data is available for the Master Data - '"+mstrdt+"'",mngdt+" - Managed Data is not available for the Master Data - '"+mstrdt+"'", "Fail");
											
											((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
											objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
											//break;
										}
						}
							
							
					}
						
				}
			}
	
	}
	*/
	public void dropdownSelectforEdit(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] Arrpr;
		Arrpr = pr.split("%");
		//Thread.sleep(5000);
		Thread.sleep(2000);
		String mngdtnd, mngdt, mstrdt, mngdts = "";
		
		
		WebElement weet = driver.findElement(By.xpath(Arrpr[0]));
		Actions act1 = new Actions(driver);
		act1.moveToElement(weet).build().perform();
		driver.findElement(By.xpath(Arrpr[1])).click();
		
		
		
		if (dt.equalsIgnoreCase("ManagedDataforEdit")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("ManagedDataforEdit");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data;
			for (int row = 3; row < rows; row++) 
			{
				
				mstrdt = sheet.getCell(0, row).getContents();
				//System.out.println(mstrdt);
				mngdtnd = sheet.getCell(3, row).getContents();
				
				if(mngdtnd.equalsIgnoreCase("Yes"))
				{
					WebElement we = driver.findElement(By.cssSelector(Arrpr[2]));
					we.click();
					System.out.println(mstrdt);
					Select se = new Select(we);
					se.selectByVisibleText(mstrdt);
					act1.click().build().perform();
					act1.sendKeys(Keys.TAB).build().perform();
					
					
					Sheet sheet1 = wb.getSheet(mstrdt);
					int rows1 = sheet1.getRows();
					for (int row1 = 0; row1 < rows1; row1++) 
					{
						System.out.println(row1);
							
						mngdt = sheet1.getCell(0, row1).getContents();
							String mngdtname = sheet1.getCell(1, row1).getContents();
								
							//System.out.println(mngdtname);
								Thread.sleep(2000);
								
								
								WebElement weall = driver.findElement(By.cssSelector(Arrpr[7]));
								
								//System.out.println("am here");
								List<WebElement> weallele = weall.findElements(By.tagName("tr"));
								int len = weallele.size();
								int flag=0;
								
								
								for(int i = 0;i<len;i++)
								{	
									
									WebElement wemngdt = driver.findElement(By.xpath("/html/body/main/div/form/div/div[2]/div[2]/table/tbody/tr['" + i + "'+1]/td[2]"));
									
									
									String wemngdtstr = wemngdt.getText(); 
									
									System.out.println(mngdt);
									System.out.println(mngdtname);
									System.out.println(wemngdtstr);
									
									System.out.println((mngdt.equalsIgnoreCase(wemngdtstr)));
									
									if(mngdt.equalsIgnoreCase(wemngdtstr))
									{
									
										WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/form/div/div[2]/div[2]/table/tbody/tr['" + i + "'+1]/td[7]/span[1]"));
									
										we5.click();
										
										Thread.sleep(2000);
										
										driver.findElement(By.xpath(Arrpr[4])).clear();
										driver.findElement(By.xpath(Arrpr[4])).sendKeys(mngdtname);		
										
										driver.findElement(By.xpath(Arrpr[6])).click();
										break;
									}
									
								}
					}
					
				}
				//break;
				
			}
		}

		
	}
	/*
	public void verifyMangedDataafterExit(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] arrPolpr;
		arrPolpr = pr.split("%");
		
		//Thread.sleep(5000);
		String mngdtnd, mngdt, mstrdt, mngdts = "";
		
		WebElement weet = driver.findElement(By.xpath("//*[@title='Admin']"));
		Thread.sleep(2000);
		Actions act1 = new Actions(driver);
		act1.moveToElement(weet).build().perform();
		driver.findElement(By.xpath("//*[@title='Manage Data']")).click();
		
		
		
		if (dt.equalsIgnoreCase("ManagedDataforEdit")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("ManagedDataforEdit");
			int rows = sheet.getRows();
			System.out.println(rows);
			String data;
			for (int row = 3; row < rows; row++) 
			{
				
				mstrdt = sheet.getCell(0, row).getContents();
				System.out.println(mstrdt);
				mngdtnd = sheet.getCell(3, row).getContents();
				
				if(mngdtnd.equalsIgnoreCase("Yes"))
				{
					WebElement we = driver.findElement(By.cssSelector("select.form-control[ng-model='ManagedDataType']"));
					we.click();
					Select se = new Select(we);
					se.selectByVisibleText(mstrdt);
					act1.click().build().perform();
					act1.sendKeys(Keys.TAB).build().perform();
					
					Sheet sheet1 = wb.getSheet(mstrdt);
					int rows1 = sheet1.getRows();
					for (int row1 = 0; row1 < rows1; row1++) 
					{
								
								mngdt = sheet1.getCell(0, row1).getContents();
								
								
								String mngdtname = sheet1.getCell(1, row1).getContents();
								
								Thread.sleep(5000);
								
								System.out.println("am here");
								
								//System.out.println(driver.findElement(By.cssSelector(arrPolpr[5])).isDisplayed());
								
								WebElement weall = driver.findElement(By.cssSelector(arrPolpr[5]));
								
								//System.out.println("am here");
								List<WebElement> weallele = weall.findElements(By.tagName("tr"));
								int len = weallele.size();
								int flag=0;
								for(int i = 0;i<len;i++)
								{	
									
									
									WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/form/div/div[2]/div[2]/table/tbody/tr['" + i + "'+1]/td[2]"));
									
									String strtxt = we5.getText();
									System.out.println(strtxt);
										System.out.println(strtxt.equalsIgnoreCase(mngdtname));
										
										
										if(strtxt.equals(mngdtname))
										{
											flag++;
											
											System.out.println("Pass");
											writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify edited Managed Data",mngdtname,mngdtname+ " - Managed Data is edited for the Master Data - '"+mstrdt+"'",strtxt+ " - Managed Data is edited for the Master Data - '"+mstrdt+"'", "Pass");
											
											
											((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
											
											//((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
											
											objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
											break;
										}
								
										
									}
										System.out.println(flag);
										if(flag==0)
										{
											System.out.println(flag);
											System.out.println("fail");
											writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify edited Managed Data",mngdtname,mngdtname+ " - Managed Data is not edited for the Master Data - '"+mstrdt+"'",mngdtname+" - Managed Data is not edited for the Master Data - '"+mstrdt+"'", "Fail");
											
											((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
											objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
											//break;
										}
						}
							
							
					}
						
				}
								
								
								
								
								
							/*	driver.findElement(By.cssSelector(Arrpr[3])).click();
								driver.findElement(By.xpath(Arrpr[4])).sendKeys(mngdt);		
								driver.findElement(By.xpath(Arrpr[6])).click();	*/
								
					
					
				
				//break;
				//row++;


		
	
	
	
	public void inFrame(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		System.out.println("am here");
		Thread.sleep(5000);
		driver.switchTo().frame(0);
		Thread.sleep(5000);
	}
	
	
	public void dropdownSelectAccType(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		
		String[] Arrpr = pr.split(DELIM);
		
		
		WebElement we = driver.findElement(By.id(Arrpr[0]));
		
		we.click();
		
		Thread.sleep(5000);
		
		Select sel = new Select(we);
		
		System.out.println(sel.getOptions());
		
		
		sel.selectByVisibleText(dt);
		
	}

public void clickByCSS(String pr, String dt)throws InterruptedException, IOException   {

	System.out.println("am here");
	
	System.out.println(driver.findElement(By.cssSelector(pr)).isDisplayed());
	
	WebElement we = driver.findElement(By.cssSelector(pr));
	WebDriverWait ww = new WebDriverWait(driver, 10);
	ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(pr)));	
	Actions act = new Actions(driver);
	
	act.click(we).build().perform();
		
	}
	public void findconfigmenu(String pr, String dt) throws InterruptedException, BiffException, IOException
	{
		Thread.sleep(5000);
		WebDriverWait ww = new WebDriverWait(driver, 10);
		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(pr)));
		
		Actions act = new Actions(driver);
		WebElement we1 = driver.findElement(By.cssSelector(pr));
		
		List<WebElement> weul = we1.findElements(By.tagName("ul"));
		for(WebElement weul1:weul)
		{
			List<WebElement> weli = weul1.findElements(By.tagName("li"));
			int lenli = weli.size();
			for(int i = 0;i<lenli;i++)
			{
				String str = weli.get(i).getText();
				if(str.equalsIgnoreCase("Configuration"))
				{
					WebElement we2 = weli.get(i);
					
					//System.out.println(we2.isDisplayed());
				
					act.moveToElement(we2).build().perform();
					
					//Thread.sleep(5000);
					
					
					
					
					WebElement wemd = driver.findElement(By.partialLinkText("Workflows"));
					
					//act.moveToElement(wemd).build().perform();
					
					act.click(wemd).build().perform();
					
					//act.sendKeys(Keys.ENTER).build().perform();
					
					//act.doubleClick(wemd).build().perform();
					
					
					
					break;
				
				}
			
			}break;
		}
	}
	public void clickWorkflow(String pr, String dt) throws InterruptedException, BiffException, IOException
	{
		
		System.out.println("am here");
		
		WebDriverWait ww = new WebDriverWait(driver, 10);
		ww.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(pr)));
		
		WebElement we1 = driver.findElement(By.cssSelector(pr));
		
		List<WebElement> weul = we1.findElements(By.tagName("ul"));
		
		
		System.out.println("1");
		for(int j = 0;j<weul.size();j++)
		{
			
			
			
			List<WebElement> weli = weul.get(j).findElements(By.tagName("li"));
			int lenli = weli.size();
			for(int i = 0;i<lenli;i++)
			{
				//System.out.println("2");
				String str = weli.get(i).getText();
				System.out.println(str);
				if(str.equalsIgnoreCase("Workflows"))
				{
					
					WebElement we2 = weli.get(i);
					Actions act = new Actions(driver);
					//act.moveToElement(we2).build().perform();
					
					//act.click(we2).build().perform();
					
					we2.findElement(By.partialLinkText("Workflows")).click();
					//act.release().build().perform();
					
					break;
				
				}
			
			}
		
		}
	}
	
	public void findusermenu(String pr, String dt) throws InterruptedException, BiffException, IOException
	{
		WebElement wemd = driver.findElement(By.partialLinkText("User"));
		
		wemd.click();
		Thread.sleep(5000);
		WebElement wemd1 = driver.findElement(By.xpath("//*[@title='Group Config']"));
	
		wemd1.click();
	}
	
	public void importusers(String pr, String dt) throws InterruptedException, BiffException, IOException
	{
	
	
		Workbook wb1 = Workbook.getWorkbook(new File(INPUTEXCEL));
		Sheet sheet1 = wb1.getSheet("Users");
		//System.out.println(sheet1.g);
		int rows1 = sheet1.getRows();
		//int cols1 = sheet1.getColumns();
		
		
			for (int row1 = 0; row1 <rows1; row1++) 
			{
				driver.findElement(By.id("ctl00_MainContent_gvUserList_ctl00_ctl02_ctl00_btnCreateUser")).click();
				Thread.sleep(8000);
				
				driver.switchTo().frame(0);
				System.out.println("Inside frame");
				System.out.println(sheet1.getCell(0, (row1+1)).getContents());
				//System.out.println(sheet1.getCell(0, 0).getContents());
				String fname = sheet1.getCell(0, (row1+1)).getContents();
				System.out.println(driver.findElement(By.id("firstName")).isDisplayed());
				
				
				
				driver.findElement(By.id("firstName")).sendKeys(fname);
				
				String lname = sheet1.getCell((1), (row1+1)).getContents();
				//System.out.println(lname);
				driver.findElement(By.id("lastname")).sendKeys(lname);
				
				
				String dname = sheet1.getCell((2), (row1+1)).getContents();
				driver.findElement(By.id("employeetitle")).sendKeys(dname);
				
				
				String depart = sheet1.getCell((3), (row1+1)).getContents();
				driver.findElement(By.id("department")).sendKeys(depart);
				
				String email = sheet1.getCell((4), (row1+1)).getContents();
				driver.findElement(By.id("email")).sendKeys(email);
				
				String uname = sheet1.getCell((5), (row1+1)).getContents();
				driver.findElement(By.id("username")).sendKeys(uname);
				
				String pwd = sheet1.getCell((6), (row1+1)).getContents();
				driver.findElement(By.id("password")).sendKeys(pwd);
				
				String confpwd = sheet1.getCell((7), (row1+1)).getContents();
				driver.findElement(By.id("cpassword")).sendKeys(confpwd);
				
				
				
				if(sheet1.getCell((8), (row1+1)).getContents().equalsIgnoreCase("Yes"))
				{
					
					driver.findElement(By.id("vieweraccess")).click();
					
				}
				
				
				String groups = sheet1.getCell((9), (row1+1)).getContents();
				String grptype = sheet1.getCell((10), (row1+1)).getContents();
				WebElement wetabs = driver.findElement(By.className("tabs"));
				List<WebElement> weuls = wetabs.findElements(By.tagName("ul"));
				for(WebElement weuls1:weuls)
				{
					List<WebElement> welis = weuls1.findElements(By.tagName("li"));
					welis.get(1).click();
					Thread.sleep(8000);
					List<WebElement> wegroups = driver.findElements(By.cssSelector("div.ng-scope[title='"+groups+"']"));
					 int wegroupslen = wegroups.size();
					for(int lisgrp=0;lisgrp<wegroupslen;lisgrp++)
					{
						System.out.println(wegroups.get(lisgrp).getText().contains(grptype));
						if(wegroups.get(lisgrp).getText().contains(grptype))
						{
							wegroups.get(lisgrp).findElement(By.xpath("//label[text()='"+grptype+"']")).click();
							
							
						}
					}
					
				}
			
		
		
		//************************************************For Save and Cancel Button
		
				driver.findElement(By.xpath("//a[text()='Save']")).click();

		//dr.findElement(By.xpath("//a[text()='Cancel']")).click();
		//System.out.println("Clicked cancel" );
		
		Thread.sleep(8000);
		driver.switchTo().defaultContent();

	
		driver.switchTo().defaultContent();
		WebElement weusrlst = driver.findElement(By.id("ctl00_MainContent_gvUserList_GridData"));
		List<WebElement> weusrlsttr = weusrlst.findElements(By.tagName("tr"));
		int weusrlsttrlen = weusrlsttr.size();
		for(int weusrlsttri=0;weusrlsttri<weusrlsttrlen;weusrlsttri++)
		{
			if(weusrlsttr.get(weusrlsttri).getText().contains(uname))
			{
				System.out.println("User "+uname+" Created Successfully");
			}
			
			break;
		}
		
		
		
		break;
		}
	
	
	}
	public void workflowCreation(String pr, String dt) throws InterruptedException, BiffException, IOException
	{
		
		
		Actions act = new Actions(driver);
		System.out.println("am here");
		WebElement we1 = driver.findElement(By.id("menu"));
		
		List<WebElement> weul = we1.findElements(By.tagName("ul"));
		for(WebElement weul1:weul)
		{
			List<WebElement> weli = weul1.findElements(By.tagName("li"));
			int lenli = weli.size();
			for(int i = 0;i<lenli;i++)
			{
				String str = weli.get(i).getText();
				if(str.equalsIgnoreCase("Configuration"))
				{
					WebElement we2 = weli.get(i);
					System.out.println(weli.get(i).getText());
					act.moveToElement(we2).build().perform();
					break;
				
				}
			
			}
		}
		WebElement wemd = driver.findElement(By.partialLinkText("Workflows"));
		wemd.click();
		Thread.sleep(5000);
	}
	
	public void userscreation(String pr, String dt) throws InterruptedException, BiffException, IOException
	{
		
		
			Actions act = new Actions(driver);
			WebElement we1 = driver.findElement(By.id("ctl00_lvMenu_menuGlams"));
			List<WebElement> weul = we1.findElements(By.tagName("ul"));
			for(WebElement weul1:weul)
			{
				List<WebElement> weli = weul1.findElements(By.tagName("li"));
				int lenli = weli.size();
				for(int i = 0;i<lenli;i++)
				{
					String str = weli.get(i).getText();
					if(str.equalsIgnoreCase("Configuration"))
					{
						WebElement we2 = weli.get(i);
					
						act.moveToElement(we2).build().perform();
						break;
					
					}
				
				}
			}
			WebElement wemd = driver.findElement(By.partialLinkText("User"));
			
			wemd.click();
			Thread.sleep(5000);
			WebElement wemd1 = driver.findElement(By.xpath("//*[@title='Group Config']"));
		
			wemd1.click();
			
			
			//dr.switchTo().frame(0);
			
			Workbook wb1 = Workbook.getWorkbook(new File(INPUTEXCEL));
			Sheet sheet1 = wb1.getSheet("Users");
			//System.out.println(sheet1.g);
			int rows1 = sheet1.getRows();
			//int cols1 = sheet1.getColumns();
			
				for (int row1 = 0; row1 <rows1; row1++) 
				{
					driver.findElement(By.id("ctl00_MainContent_gvUserList_ctl00_ctl02_ctl00_btnCreateUser")).click();
					Thread.sleep(8000);
					
					driver.switchTo().frame(0);
					System.out.println("Inside frame");
					System.out.println(sheet1.getCell(0, (row1+1)).getContents());
					//System.out.println(sheet1.getCell(0, 0).getContents());
					String fname = sheet1.getCell(0, (row1+1)).getContents();
					System.out.println(driver.findElement(By.id("firstName")).isDisplayed());
					
					
					
					driver.findElement(By.id("firstName")).sendKeys(fname);
					
					String lname = sheet1.getCell((1), (row1+1)).getContents();
					//System.out.println(lname);
					driver.findElement(By.id("lastname")).sendKeys(lname);
					
					
					String dname = sheet1.getCell((2), (row1+1)).getContents();
					driver.findElement(By.id("employeetitle")).sendKeys(dname);
					
					
					String depart = sheet1.getCell((3), (row1+1)).getContents();
					driver.findElement(By.id("department")).sendKeys(depart);
					
					String email = sheet1.getCell((4), (row1+1)).getContents();
					driver.findElement(By.id("email")).sendKeys(email);
					
					String uname = sheet1.getCell((5), (row1+1)).getContents();
					driver.findElement(By.id("username")).sendKeys(uname);
					
					String pwd = sheet1.getCell((6), (row1+1)).getContents();
					driver.findElement(By.id("password")).sendKeys(pwd);
					
					String confpwd = sheet1.getCell((7), (row1+1)).getContents();
					driver.findElement(By.id("cpassword")).sendKeys(confpwd);
					
					
					
					if(sheet1.getCell((8), (row1+1)).getContents().equalsIgnoreCase("Yes"))
					{
						
						driver.findElement(By.id("vieweraccess")).click();
						
					}
					
					
					String groups = sheet1.getCell((9), (row1+1)).getContents();
					String grptype = sheet1.getCell((10), (row1+1)).getContents();
					WebElement wetabs = driver.findElement(By.className("tabs"));
					List<WebElement> weuls = wetabs.findElements(By.tagName("ul"));
					for(WebElement weuls1:weuls)
					{
						List<WebElement> welis = weuls1.findElements(By.tagName("li"));
						welis.get(1).click();
						Thread.sleep(8000);
						List<WebElement> wegroups = driver.findElements(By.cssSelector("div.ng-scope[title='"+groups+"']"));
						 int wegroupslen = wegroups.size();
						for(int lisgrp=0;lisgrp<wegroupslen;lisgrp++)
						{
							System.out.println(wegroups.get(lisgrp).getText().contains(grptype));
							if(wegroups.get(lisgrp).getText().contains(grptype))
							{
								wegroups.get(lisgrp).findElement(By.xpath("//label[text()='"+grptype+"']")).click();
								
								
							}
						}
						
					}
				
			
			
			//************************************************For Save and Cancel Button
			
					driver.findElement(By.xpath("//a[text()='Save']")).click();

			//dr.findElement(By.xpath("//a[text()='Cancel']")).click();
			//System.out.println("Clicked cancel" );
			
			Thread.sleep(8000);
			driver.switchTo().defaultContent();

		
			driver.switchTo().defaultContent();
			WebElement weusrlst = driver.findElement(By.id("ctl00_MainContent_gvUserList_GridData"));
			List<WebElement> weusrlsttr = weusrlst.findElements(By.tagName("tr"));
			int weusrlsttrlen = weusrlsttr.size();
			for(int weusrlsttri=0;weusrlsttri<weusrlsttrlen;weusrlsttri++)
			{
				if(weusrlsttr.get(weusrlsttri).getText().contains(uname))
				{
					System.out.println("User "+uname+" Created Successfully");
				}
				
				break;
			}
			
			
			
			break;
			}
		
		
	}
		
		
		
		
	public void moveByXpath(String pr, String dt)throws InterruptedException, IOException   {

		
		
		WebElement we = driver.findElement(By.xpath(pr));
		WebDriverWait ww = new WebDriverWait(driver, 10);
		ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
		
		Actions act = new Actions(driver);
		act.moveToElement(we).build().perform();
		//Thread.sleep(5000);
		
	}
	
	
	
	
	
public void clickByXpath(String pr, String dt)throws InterruptedException, IOException   {

	
	
	WebElement we = driver.findElement(By.xpath(pr));
	WebDriverWait ww = new WebDriverWait(driver, 10);
	
	ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));	
	Actions act = new Actions(driver);
	act.click(we).build().perform();
		
	}
	
	
	
	/*
	@SuppressWarnings("unused")
	public void verifyMasterDt(String pr, String dt)throws InterruptedException, IOException, BiffException   {
		
		
		String[] arrPolpr = pr.split(DELIM);
		

		
		String str ="";
		int i,j, row;
		if (dt.equalsIgnoreCase("MasterData")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("MasterData");
			int rows = sheet.getRows();
			//System.out.println(rows);
			String data="";
			for (row = 3; row < rows; row++) 
			{
				data = sheet.getCell(0, row).getContents();
				
				System.out.println(data);
				
				WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
				
				//System.out.println("am here");
				List<WebElement> we4 = we3.findElements(By.tagName("tr"));
				int len = we4.size();
				int flag=0;
				for(i = 1;i<len;i++)
				{	
					
					
					WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]"));
					str = we5.getText();
					System.out.println(str);
						System.out.println(str.equalsIgnoreCase(data));
						
						
						if(str.equals(data))
						{
							flag++;
							
							System.out.println("Pass");
							writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",data,data+ " - Master Data is available",str+ " - Master Data is available", "Pass");
							
							
							((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
							
							//((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
							
							objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
							break;
						}
				
						/*if(we5.isDisplayed())
						{
							System.out.println("fail");
						}*/
/*
				}
						System.out.println(flag);
						if(flag==0)
						{
							System.out.println(flag);
							System.out.println("fail");
							writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",data,data+ " - Master Data is available",data+" - Master Data is not available", "Fail");
							
							((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
							objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
							//break;
						}
					}
			
			
		}
		
	}
	*/
	
public void verifyMasterDtEdit(String pr, String dt)throws InterruptedException, IOException, BiffException   {
		
		
	/*	String[] arrPolpr = pr.split(DELIM);
		

		
		String str ="";
		int i,j, row;
		if (dt.equalsIgnoreCase("MasterData")) 
		{
			
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("MasterData");
			int rows = sheet.getRows();
			//System.out.println(rows);
			String data="";
			for (row = 3; row < rows; row++) 
			{
				data = sheet.getCell(0, row).getContents();
				
				System.out.println(data);
				
				WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
				
				//System.out.println("am here");
				List<WebElement> we4 = we3.findElements(By.tagName("tr"));
				int len = we4.size();
				int flag=0;
				for(i = 1;i<len;i++)
				{	
					
					
					WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]"));
					str = we5.getText();
					System.out.println(str);
						System.out.println(str.equalsIgnoreCase(data));
						
						
						if(str.equals(data))
						{
							flag++;
							
							System.out.println("Pass");
							writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",data,data+ " - Master Data is available",str+ " - Master Data is available", "Pass");
							
							
							((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
							
							//((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
							
							objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
							break;
						}
				
						/*if(we5.isDisplayed())
						{
							System.out.println("fail");
						}*/
				}
/*
						System.out.println(flag);
						if(flag==0)
						{
							System.out.println(flag);
							System.out.println("fail");
							writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",data,data+ " - Master Data is not available",data+" - Master Data is not available", "Fail");
							
							((JavascriptExecutor) driver).executeScript("window.scrollTo(0, document.body.scrollHeight)");
							objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
							//break;
						}
					}
			
			
		}
		
	}*/
	
public void usersCreation(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		String[] arrPolpr = pr.split("%");
		
		if(dt.equalsIgnoreCase("UsersList"))
		{
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("UsersList");
			int rows = sheet.getRows();
			String FirstName,LastName, DisplayName, Department, Email, AccountType, UserName, Password, ConfirmPassword, ViewerAccess="";
			
			for (int row = 1; row < rows; row++) 
			{
				Thread.sleep(5000);
				//driver.switchTo().defaultContent();
				driver.findElement(By.className("ng-isolate-scope")).click();
				
				driver.switchTo().frame(0);
				
				FirstName = sheet.getCell(0,row).getContents();
				LastName = sheet.getCell(1,row).getContents();
				DisplayName = sheet.getCell(2,row).getContents();
				Department = sheet.getCell(3,row).getContents();
				Email = sheet.getCell(4,row).getContents();
				AccountType = sheet.getCell(5,row).getContents();
				UserName = sheet.getCell(6,row).getContents();
				Password = sheet.getCell(7,row).getContents();
				ConfirmPassword = sheet.getCell(8,row).getContents();
				ViewerAccess = sheet.getCell(9,row).getContents();
				
				
				driver.findElement(By.id(arrPolpr[0])).sendKeys(FirstName);
							
				driver.findElement(By.cssSelector(arrPolpr[1])).sendKeys(LastName);
				
				
				driver.findElement(By.xpath(arrPolpr[2])).sendKeys(DisplayName);
				
				driver.findElement(By.xpath(arrPolpr[3])).sendKeys(Department);
				
				driver.findElement(By.xpath(arrPolpr[4])).sendKeys(Email);
				
				WebElement we = driver.findElement(By.xpath(arrPolpr[5]));
				we.click();
				
				Select sel = new Select(we);
				sel.selectByVisibleText(AccountType);
				
				
				//sel.selectByIndex(0);
				
				Actions act1 = new Actions(driver);
				act1.click(we).build().perform();
				
				act1.sendKeys(Keys.ARROW_DOWN).build().perform();
				act1.sendKeys(Keys.ENTER).build().perform();
				we.click();
				
				
				System.out.println("am here");
				Thread.sleep(5000);
				System.out.println(driver.findElement(By.xpath(arrPolpr[6])).isDisplayed());
				driver.findElement(By.xpath(arrPolpr[6])).sendKeys(UserName);
				
				
				driver.findElement(By.xpath(arrPolpr[7])).sendKeys(Password);
				Thread.sleep(5000);
				driver.findElement(By.xpath(arrPolpr[8])).sendKeys(ConfirmPassword);
				
				if(ViewerAccess.equals("Yes"))
				{
					driver.findElement(By.xpath(arrPolpr[9])).click();
				}
				 
				System.out.println(arrPolpr[10]);
				System.out.println(driver.findElement(By.xpath(arrPolpr[10])).isDisplayed());
				
				driver.findElement(By.xpath(arrPolpr[10])).click();
				
				
//				 System.out.println(driver.findElement(By.xpath(arrPolpr[11])).isDisplayed());
//				 WebElement weclose = driver.findElement(By.xpath(arrPolpr[11]));
//				 Thread.sleep(10000);
//				 weclose.click();

				/*
				WebDriverWait wa = new WebDriverWait(driver, 20);
				wa.until(ExpectedConditions.elementToBeClickable(weclose));*/
				
				
		
			}
		}
			
		
	}
	
	/*
public void verifyUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
			
	
			String[] arrPolpr = pr.split(DELIM);
			String str ="";
			int i,j, row;
			if (dt.equalsIgnoreCase("UsersList")) 
			{
				
					Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
					Sheet sheet = wb.getSheet("UsersList");
					int rows = sheet.getRows();
					//System.out.println(rows);
					String data="";
					for (row = 1; row < rows; row++) 
					{
						data = sheet.getCell(0, row).getContents();
						
						System.out.println(data);
						
						WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
						
						//System.out.println("am here");
						List<WebElement> we4 = we3.findElements(By.tagName("tr"));
						int len = we4.size();
						int flag=0;
						for(i = 1;i<len;i++)
						{	
							
							WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]/a"));
							str = we5.getText();
							System.out.println(str);
							System.out.println(str.equalsIgnoreCase(data));
							if(str.equals(data))
							{
								flag++;
								System.out.println("Pass");
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Users",data,data+ " - User is available",str+ " - User is available", "Pass");
								
								((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
								objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
								break;
							}
						}
			
						if(flag==0)
						{
							for(i = 1;i<len;i++)
							{	
								
								WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]/a"));
							
								System.out.println(flag);
								
								
								System.out.println("fail");
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify User",data,data+ " - User is not available",data+" - User is not available", "Fail");
								((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
								objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
								break;
							}
						}
					}
			}
}




public void selectGroupsandPermissions(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	
	String[] arrPolpr = pr.split(DELIM);
	String str ="";
	int i,j,k, row, flagdata=1, flaggroups=0, flagstr1;
	
	if (dt.equalsIgnoreCase("SelectUser")) 
	{
		
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("SelectUser");
			int rows = sheet.getRows();
			int cols1 = sheet.getColumns();
			System.out.println(cols1);
			System.out.println(rows);
			String data,groups, str1="";
			
			for(int col=0;col<cols1;col++)
			{
				/*for (row=1;row<rows;row++)
				{*/
				/*
					data = sheet.getCell(col, flagdata).getContents();
					groups = sheet.getCell(col+1, flaggroups).getContents();
					str1= sheet.getCell(col+1, flagdata).getContents();
				
					System.out.println(data);
					System.out.println(groups);
					System.out.println(str1);
				
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
				
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(i = 1;i<=len;i++)
					{	
					
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equalsIgnoreCase(data))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]/a")).click();
							driver.switchTo().frame(0);
							driver.findElement(By.xpath(arrPolpr[5])).click();
					
						
						
							Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
							Sheet sheet1 = wb1.getSheet(groups);
							int colssh1 = sheet1.getColumns();
							for(k=0;k<colssh1;k++)
							{
								String sublist = sheet1.getCell(k,0).getContents();
								String sublistneed = sheet1.getCell(k,1).getContents();
								List<WebElement> we = driver.findElements(By.cssSelector(arrPolpr[6]));
								int lisize=we.size();
								for(j=0;j<lisize;j++)
								{
									String str2 = driver.findElement(By.xpath("/html/body/form/div[1]/div[2]/div[2]/div/div[2]/div['"+j+"'+1]/div/div[1]")).getText();
							
									if((groups.equalsIgnoreCase(str2))&&((str1.equalsIgnoreCase("Yes")))) 
									{
										System.out.println(sublist);
										System.out.println(sublistneed);
									}
								}
							}
							
							//break;
						} 
					
						
						flagdata++;
						flaggroups++;
					}
					System.out.println("am here2");
					
				
					
					
				}
			
		}
	
}
	


*/



public void selectEditRoles(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	
	String[] arrPolpr = pr.split(DELIM);
	String str ="";
	int i,j,k,l,row;
	String strtemp="";
	if (dt.equals("UsersforEditRoles")) 
	{
		
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("UsersforEditRoles");
			int rows = sheet.getRows();
			int cols = sheet.getColumns();
			int flag=1, flag1=1, flag2=1;
			
			String users, temps, needs="";
			for (i=1;i<rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					System.out.println(users);
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							driver.findElement(By.cssSelector(arrPolpr[5])).click();
							Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
							Sheet sheet1 = wb1.getSheet(users);
							int cols1 = sheet1.getColumns();
							for (j = 0; j < cols1; j++)
							{
								temps = sheet1.getCell(j,0).getContents();
								needs = sheet1.getCell(j,1).getContents();
								System.out.println("from Excel = --- " +temps);
								System.out.println("from Excel = --- " +needs);
								List<WebElement> wetemps = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Role-Template.user-management.ng-scope div.clearfix div.row div.column.span-8 div.option.ng-scope label"));
								
								for(l=0;l<wetemps.size();l++)
								{
									strtemp = wetemps.get(l).getText();
									System.out.println("from app = --- " +wetemps.get(l).getText());
									//System.out.println(temps);
									
									
									//System.out.println(strtemp.equalsIgnoreCase(temps));
									
									//System.out.println(l);
									
									//System.out.println(needs.equalsIgnoreCase("Yes"));
									
									if(temps.equalsIgnoreCase(strtemp) && (needs.equalsIgnoreCase("Yes")))
										{
											System.out.println("me here");
											String strclicked = wetemps.get(l).getText();
											wetemps.get(l).click();
											
											//System.out.println(driver.findElement(By.xpath("/html/body/div[109]/div/div[2]/div/div[2]/div/div[2]/div")).isDisplayed());
											System.out.println(driver.findElement(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Role-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer span.selectedlist.ng-scope.ng-binding")).isDisplayed());
											List<WebElement> wetemps1 = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Role-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer span.selectedlist.ng-scope.ng-binding"));
											System.out.println(wetemps1.size());
											for(int p=0;p<wetemps1.size();p++)
											{
												System.out.println("Selected list - "+wetemps1.get(p).getText());
												String added = wetemps1.get(p).getText();
												System.out.println(strclicked.equalsIgnoreCase(added));
												if(strclicked.equalsIgnoreCase(added))
												{
													System.out
															.println("errror is here");
													writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Selected Role's of User - " +users+ "and Role of - " +strclicked,strclicked,strclicked+ " - Role is created or not for the User - "+users,added+ " - Role is created for the User - "+users,"Pass");
													break;
												}
											}
											
												//System.out.println(strclicked.equalsIgnoreCase(added))
										}
								}
							}
							driver.findElement(By.xpath("/html/body/div[55]/div/div[3]/a[2]")).click();
							driver.findElement(By.cssSelector("a.button[ng-click='close()']")).click();
							
							driver.switchTo().defaultContent();
							break;
							
						}
					}
					
					
				}
			
		}
	
	
	}
	
public void verifyEditedRoles(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	System.out.println("am here");
	
	
	String[] arrPolpr = pr.split(DELIM);
	String str ="";
	int i,j,k,l,row;
	String strtemp="";
	if (dt.equals("EditEmailTemplateUsers")) 
	{
		
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("EditEmailTemplateUsers");
			int rows = sheet.getRows();
			int cols = sheet.getColumns();
			int flag=1, flag1=1, flag2=1;
			
			String users, temps, needs="";
			for (i=1;i<=rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					System.out.println(users);
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							driver.findElement(By.xpath(arrPolpr[5])).click();
							driver.findElement(By.cssSelector(arrPolpr[6])).click();
							Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
							Sheet sheet1 = wb1.getSheet(users);
							int cols1 = sheet1.getColumns();
							for (j = 0; j < cols1; j++)
							{
								temps = sheet1.getCell(j,0).getContents();
								needs = sheet1.getCell(j,1).getContents();
								System.out.println("from Excel = --- " +temps);
								System.out.println("from Excel = --- " +needs);
								
								List<WebElement> wetemps = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Email-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer"));
								
								System.out.println(wetemps.size());
								for(l=0;l<wetemps.size();l++)
								{
									strtemp = wetemps.get(l).getText();
									System.out.println("from app = --- " +wetemps.get(l).getText());
									System.out.println(temps);
									
									
									if(temps.equalsIgnoreCase(strtemp)&&(needs.equalsIgnoreCase("Yes")))
										{
											System.out.println("me here");
											System.out.println(wetemps.get(l).getText());
											wetemps.get(l).click();
											String strclicked = wetemps.get(l).getText();
											System.out.println(driver.findElement(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Role-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer span.selectedlist.ng-scope.ng-binding")).isDisplayed());
											List<WebElement> wetemps1 = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Role-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer span.selectedlist.ng-scope.ng-binding"));
											System.out.println(wetemps1.size());
											for(int p=0;p<wetemps1.size();p++)
											{
												System.out.println("Selected list - "+wetemps1.get(p).getText());
												String added = wetemps1.get(p).getText();
												System.out.println(strclicked.equalsIgnoreCase(added));
												if(strclicked.equalsIgnoreCase(added))
												{
													writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Selected Role's of User - " +users+ "and Role of - " +strclicked,strclicked,strclicked+ " - Role is created or not for the User - "+users,added+ " - Role is created for the User - "+users,"Pass");
													break;
												}
											}
											
											
										}
									
									
								}
							}
							driver.findElement(By.xpath("/html/body/div[55]/div/div[3]/a[2]")).click();
							driver.findElement(By.cssSelector("a.button[ng-click='close()']")).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
				}
		}
	}


public void selectEmailTemplateUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	
	String[] arrPolpr = pr.split(DELIM);
	String str ="";
	int i,j,k,l,row;
	String strtemp="";
	if (dt.equals("EditEmailTemplateUsers")) 
	{
		
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("EditEmailTemplateUsers");
			int rows = sheet.getRows();
			//int cols = sheet.getColumns();
			//int flag=1, flag1=1, flag2=1;
			
			String users, temps, needs="";
			for (i=1;i<rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					System.out.println(users);
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							driver.findElement(By.xpath(arrPolpr[5])).click();
							driver.findElement(By.cssSelector(arrPolpr[6])).click();
							Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
							Sheet sheet1 = wb1.getSheet(users);
							int cols1 = sheet1.getColumns();
							for (j = 0; j < cols1; j++)
							{
								temps = sheet1.getCell(j,0).getContents();
								needs = sheet1.getCell(j,1).getContents();
								System.out.println("from Excel = --- " +temps);
								System.out.println("from Excel = --- " +needs);
								
								List<WebElement> wetemps = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Email-Template.user-management.ng-scope div.clearfix div.row div.column.span-8 div.option.ng-scope label.ng-binding"));
								
								//System.out.println(wetemps.size());
								for(l=0;l<wetemps.size();l++)
								{
									strtemp = wetemps.get(l).getText();
									System.out.println("from app = --- " +wetemps.get(l).getText());
									
									if(temps.equalsIgnoreCase(strtemp)&&(needs.equalsIgnoreCase("Yes")))
										{
											System.out.println("me here");
											String strclicked = wetemps.get(l).getText();
											wetemps.get(l).click();
											
											System.out.println(driver.findElement(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Email-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer span.selectedlist.ng-scope.ng-binding")).isDisplayed());
											List<WebElement> wetemps1 = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Email-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer span.selectedlist.ng-scope.ng-binding"));
											System.out.println(wetemps1.size());
											for(int p=0;p<wetemps1.size();p++)
											{
												System.out.println("Selected list - "+wetemps1.get(p).getText());
												String added = wetemps1.get(p).getText();
												System.out.println(strclicked.equalsIgnoreCase(added));
												if(strclicked.equalsIgnoreCase(added))
												{
													System.out.println("errror is here");
													writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Edit Email Template of User - " +users+ " and Template of - " +strclicked,strclicked,strclicked+ " - Temaplate is created or not for the User - "+users,added+ " - Template is created for the User - "+users,"Pass");
													break;
												}
											}
											
										}
									
									
								}
							}
							driver.findElement(By.xpath("/html/body/div[55]/div/div[3]/a[2]")).click();
							driver.findElement(By.cssSelector("a.button[ng-click='close()']")).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
				}
		}
	}




public void selectChangePwdUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	//String[] arrPoldt = dt.split("%");
	String str ="";
	int i,j,k,l,row;
	String cnfmpwd="";
	if (dt.equals("UsersforChangepwd")) 
	{
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("UsersforChangepwd");
			int rows = sheet.getRows();
			
			String users, pwds = "";
			for (i=1;i<rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					pwds = sheet.getCell(1,i).getContents();
					cnfmpwd = sheet.getCell(2,i).getContents();
					System.out.println(users);
					System.out.println(pwds);
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							driver.findElement(By.cssSelector(arrPolpr[5])).click();
							driver.findElement(By.cssSelector(arrPolpr[6])).sendKeys(pwds);
							driver.findElement(By.xpath(arrPolpr[7])).sendKeys(cnfmpwd);
							driver.findElement(By.linkText("Cancel")).click();
							driver.findElement(By.partialLinkText(arrPolpr[8])).click(); // for Save button
							String strexp = driver.findElement(By.cssSelector(arrPolpr[9])).getText();
							if(strexp.equals("User updated successfully"))
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Password has been changed for the User - " +users,users,users+ " - Password has been changed for the User - "+users,"Password has been changed for the User - "+users,"Pass");
							}
							else
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Password has been changed for the User - " +users,users,users+ " - Password has been changed for the User - "+users,"Password has not been changed for the User - "+users,"Pass");
							}
							Thread.sleep(5000);
							driver.findElement(By.cssSelector(arrPolpr[10])).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
					
					
			}
	}
}


public void selectResetPwdUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	//String[] arrPoldt = dt.split("%");
	String str ="";
	int i,j,k,l,row;
	String cnfmpwd="";
	if (dt.equals("UsersforResetpwd")) 
	{
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("UsersforResetpwd");
			int rows = sheet.getRows();
			
			String users= "";
			for (i=1;i<rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							driver.findElement(By.cssSelector(arrPolpr[5])).click();
							
							driver.findElement(By.linkText("Reset Password")).click();
							
//							driver.findElement(By.linkText("Cancel")).click();
//							driver.findElement(By.partialLinkText(arrPolpr[8])).click(); // for Save button
							String strexp = driver.findElement(By.cssSelector(arrPolpr[9])).getText();
							if(strexp.equals("User updated successfully"))
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Password has been Reset for the User - " +users,users,users+ " - Password has been Reset for the User - "+users,"Password has been Reset for the User - "+users,"Pass");
							}
							else
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Password has been Reset for the User - " +users,users,users+ " - Password has been Reset for the User - "+users,"Password has not been Reset for the User - "+users,"Fail");
							}
							Thread.sleep(5000);
							driver.findElement(By.cssSelector(arrPolpr[10])).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
					
					
			}
	}
}




public void selectLockUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	//String[] arrPoldt = dt.split("%");
	String str ="";
	int i,j,k,l,row;
	String cnfmpwd="";
	if (dt.equals("LockUnLockUsr")) 
	{
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("LockUnLockUsr");
			int rows = sheet.getRows();
			
			String users= "";
			for (i=1;i<rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							//driver.findElement(By.cssSelector(arrPolpr[5])).click();
							
							driver.findElement(By.linkText("Lock User")).click();
							
//							driver.findElement(By.linkText("Cancel")).click();
//							driver.findElement(By.partialLinkText(arrPolpr[8])).click(); // for Save button
							String strexp = driver.findElement(By.cssSelector(arrPolpr[9])).getText();
							if(strexp.equals("User locked successfully"))
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify User has been Locked - " +users,users,users+ " - User has been Locked",users+ " - User has been Locked","Pass");
							}
							else
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify User has been Locked - " +users,users,users+ " - User has been Locked",users+ " - User has not been Locked","Fail");
							}
							Thread.sleep(5000);
							driver.findElement(By.cssSelector(arrPolpr[10])).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
					
					
			}
	}
}





public void selectUnLockUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	//String[] arrPoldt = dt.split("%");
	String str ="";
	int i,j,k,l,row;
	String cnfmpwd="";
	if (dt.equals("LockUnLockUsr")) 
	{
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("LockUnLockUsr");
			int rows = sheet.getRows();
			
			String users= "";
			for (i=1;i<rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							//driver.findElement(By.cssSelector(arrPolpr[5])).click();
							
							driver.findElement(By.linkText("Unlock User")).click();
							
//							driver.findElement(By.linkText("Cancel")).click();
//							driver.findElement(By.partialLinkText(arrPolpr[8])).click(); // for Save button
							String strexp = driver.findElement(By.cssSelector(arrPolpr[9])).getText();
							if(strexp.equals("User unlocked successfully"))
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify User has been UnLocked - " +users,users,users+ " - User has been UnLocked",users+ " - User has been UnLocked","Pass");
							}
							else
							{
								writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify User has been UnLocked - " +users,users,users+ " - User has been UnLocked",users+ " - User has not been UnLocked","Fail");
							}
							Thread.sleep(5000);
							driver.findElement(By.cssSelector(arrPolpr[10])).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
					
					
			}
	}
}




public void verifyEmailTemplateUsers(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	
	String[] arrPolpr = pr.split(DELIM);
	String str ="";
	int i,j,k,l,row;
	String strtemp="";
	if (dt.equals("EditEmailTemplateUsers")) 
	{
		
			Workbook wb = Workbook.getWorkbook(new File(DRIVEREXCEL));
			Sheet sheet = wb.getSheet("EditEmailTemplateUsers");
			int rows = sheet.getRows();
			//int cols = sheet.getColumns();
			//int flag=1, flag1=1, flag2=1;
			
			String users, temps, needs="";
			for (i=1;i<=rows;i++)
			{
				
					users = sheet.getCell(0,i).getContents();
					System.out.println(users);
					
					WebElement we3 = driver.findElement(By.className(arrPolpr[4]));
					
					List<WebElement> we4 = we3.findElements(By.tagName("tr"));
					int len = we4.size();
				
					for(k=1;k<=len;k++)
					{	
						WebElement we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a"));
						str = we5.getText();
						if(str.equals(users))
						{
							driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + k + "'+1]/td[1]/a")).click();
							Thread.sleep(5000);
							driver.switchTo().frame(0);
							driver.findElement(By.xpath(arrPolpr[5])).click();
							driver.findElement(By.cssSelector(arrPolpr[6])).click();
							Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
							Sheet sheet1 = wb1.getSheet(users);
							int cols1 = sheet1.getColumns();
							for (j = 0; j < cols1; j++)
							{
								temps = sheet1.getCell(j,0).getContents();
								needs = sheet1.getCell(j,1).getContents();
								System.out.println("from Excel = --- " +temps);
								System.out.println("from Excel = --- " +needs);
								
								List<WebElement> wetemps = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Email-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer"));
								
								System.out.println(wetemps.size());
								for(l=0;l<wetemps.size();l++)
								{
									strtemp = wetemps.get(l).getText();
									System.out.println("from app = --- " +wetemps.get(l).getText());
									System.out.println(temps);
									System.out.println(strtemp.equals(temps));
									System.out.println(l);
									System.out.println(needs.equalsIgnoreCase("Yes"));
									if(temps.equalsIgnoreCase(strtemp)&&(needs.equalsIgnoreCase("Yes")))
										{
											System.out.println("me here");
											System.out.println(wetemps.get(l).getText());
											wetemps.get(l).click();
											String clickedtext = wetemps.get(l).getText();
											System.out.println(clickedtext);
											
											List<WebElement> wetemps1 = driver.findElements(By.cssSelector("html body div.popup-wrapper div.popup.ui-draggable div.popup-content div#Email-Template.user-management.ng-scope div.clearfix div.row div.column.span-4 div.selectedListContainer"));
											
											System.out.println(wetemps1.size());
											
											
										}
									
									
								}
							}
							driver.findElement(By.xpath("/html/body/div[55]/div/div[3]/a[2]")).click();
							driver.findElement(By.cssSelector("a.button[ng-click='close()']")).click();
							driver.switchTo().defaultContent();
							break;
						}
					}
				}
		}
	}






public void selectGroupsandPermTab(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split(DELIM);
	String str ="";
	int i,j, row;
	driver.switchTo().frame(0);
	
	
	driver.findElement(By.xpath(arrPolpr[5])).click();
	List<WebElement> we = driver.findElements(By.cssSelector(arrPolpr[6]));
	int lisize=we.size();
	for( i=0;i<lisize;i++)
	{
		str = we.get(i).getAttribute("title");
		System.out.println(str);
	}
	
	
	
	
	
}


/*public void verifyManagedDt(String pr, String dt)throws InterruptedException, IOException   {
		
	String[] arrPolpr = pr.split(DELIM);
	
	
	WebElement we3 = driver.findElement(By.cssSelector("div.ng-valid[ng-model='SelectedList']"));
	
	
		
		
	}*/
	
	
public void selectLayout(String pr, String dt)throws InterruptedException, IOException, BiffException   {
		
		
		
		
		String[] arrPolpr = pr.split(DELIM);
		
		
		WebElement we = driver.findElement(By.className(arrPolpr[4]));
		
		List<WebElement> wetr = we.findElements(By.tagName("tr")) ;
		int flag=0;
		for(int i=1;i<wetr.size();i++)
		{
			
			WebElement wetd = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['"+i+"'+1]/td[1]/a"));
			
			//System.out.println("am here");
			
			String str = wetd.getText();
				
			if(str.equals(dt))
				
			{
				//System.out.println(str);
				//System.out.println(dt);
				//wetd.findElement(By.).click();
				
				driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['"+i+"'+1]/td[1]/a")).click();
				
				
				flag++;
				//System.out.println("pass");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Layout",dt,dt+ " - Layout is existed",str+ " - Layout is existed", "Pass");
				
				
			}
			
				
		}
		
		if(flag==0)
		{
			
			//System.out.println("Fail");
			writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",dt,dt+ " - Layout is not existed",dt+ " - Layout is not existed", "Fail");
			
			
		}
		
	}



public void dragTextBox(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	
	if(dt.equalsIgnoreCase("TextBoxes"))
	{
		
		Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
		Sheet sheet1 = wb1.getSheet("TextBoxes");
		int rows1 = sheet1.getRows();
		
		int cols1 = sheet1.getColumns();
		String data="";
		String Label, WaterMarkText,Name,Value,Maxlength,size,ValidationRequired,PostBack,ItemInfo,Visible,ReadOnly="";
		int count = 1;
		for (int row1 = 1; row1 < rows1; row1++) 
		{
			
			Label = sheet1.getCell(0, row1).getContents();
			WaterMarkText = sheet1.getCell(1, row1).getContents();
			Name = sheet1.getCell(2, row1).getContents();
			Value = sheet1.getCell(3, row1).getContents();
			Maxlength = sheet1.getCell(4, row1).getContents();
			size = sheet1.getCell(5, row1).getContents();
			ValidationRequired = sheet1.getCell(6, row1).getContents();
			PostBack = sheet1.getCell(7, row1).getContents();
			ItemInfo = sheet1.getCell(8, row1).getContents();
			Visible = sheet1.getCell(9, row1).getContents();
			ReadOnly = sheet1.getCell(10, row1).getContents();
			
			//System.out.println("am here	");
			
			System.out.println(Label);
			
			WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
			WebElement we1 = driver.findElement(By.cssSelector(arrPolpr[1]));
			Actions act = new Actions(driver);
			act.clickAndHold(we).build().perform();
			act.release(we1).build().perform();
			Thread.sleep(5000);
			
			driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div[2]/div/div/div/div/div/div[2]/div["+count+"]")).click();
			
			driver.findElement(By.cssSelector(arrPolpr[3])).clear();
			driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(Label);
				
			driver.findElement(By.cssSelector(arrPolpr[4])).clear();
			driver.findElement(By.cssSelector(arrPolpr[4])).sendKeys(WaterMarkText);
					
			driver.findElement(By.cssSelector(arrPolpr[5])).clear();
			driver.findElement(By.cssSelector(arrPolpr[5])).sendKeys(Name);
					
			driver.findElement(By.cssSelector(arrPolpr[6])).clear();
			driver.findElement(By.cssSelector(arrPolpr[6])).sendKeys(Value);
					
			driver.findElement(By.cssSelector(arrPolpr[7])).clear();
			driver.findElement(By.cssSelector(arrPolpr[7])).sendKeys(Maxlength);
				
			driver.findElement(By.cssSelector(arrPolpr[8])).clear();
			driver.findElement(By.cssSelector(arrPolpr[8])).sendKeys(size);
				
			if(ValidationRequired.equalsIgnoreCase("Yes"))
			{
				System.out.println(arrPolpr[9]);
				driver.findElement(By.cssSelector(arrPolpr[9])).click();
			}
			if(PostBack.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[10])).click();
			}
			if(ItemInfo.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[11])).click();
			}
			if(Visible.equalsIgnoreCase("No"))
			{
				driver.findElement(By.cssSelector(arrPolpr[12])).click();
			}
			if(ReadOnly.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[13])).click();
			}
		
			 count++;
		}		
	}
	
}

public void dragTextBoxinCol(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	
	if(dt.equalsIgnoreCase("TextBoxes"))
	{
		
		Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
		Sheet sheet1 = wb1.getSheet("TextBoxes");
		int rows1 = sheet1.getRows();
		
		int cols1 = sheet1.getColumns();
		System.out.println(cols1);
		
		System.out.println(rows1);
		String data="";
		String Label, WaterMarkText,Name,Value,Maxlength,size,ValidationRequired,PostBack,ItemInfo,Visible,ReadOnly="";
		int count = 1;
		for (int row1 = 1; row1 < rows1; row1++) 
		{
			
			Label = sheet1.getCell(0, row1).getContents();
			WaterMarkText = sheet1.getCell(1, row1).getContents();
			Name = sheet1.getCell(2, row1).getContents();
			Value = sheet1.getCell(3, row1).getContents();
			Maxlength = sheet1.getCell(4, row1).getContents();
			size = sheet1.getCell(5, row1).getContents();
			ValidationRequired = sheet1.getCell(6, row1).getContents();
			PostBack = sheet1.getCell(7, row1).getContents();
			ItemInfo = sheet1.getCell(8, row1).getContents();
			Visible = sheet1.getCell(9, row1).getContents();
			ReadOnly = sheet1.getCell(10, row1).getContents();
			
			//System.out.println("am here	");
			
			WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
			WebElement we1 = driver.findElement(By.cssSelector(arrPolpr[1]));
			Actions act = new Actions(driver);
			act.clickAndHold(we).build().perform();
			act.release(we1).build().perform();
			Thread.sleep(5000);
			
			driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div[2]/div/div/div/div/div/div[2]/div/div/div[1]/div/div["+count+"]")).click();
					  
			driver.findElement(By.cssSelector(arrPolpr[3])).clear();
			driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(Label);
				
			driver.findElement(By.cssSelector(arrPolpr[4])).clear();
			driver.findElement(By.cssSelector(arrPolpr[4])).sendKeys(WaterMarkText);
					
			driver.findElement(By.cssSelector(arrPolpr[5])).clear();
			driver.findElement(By.cssSelector(arrPolpr[5])).sendKeys(Name);
					
			driver.findElement(By.cssSelector(arrPolpr[6])).clear();
			driver.findElement(By.cssSelector(arrPolpr[6])).sendKeys(Value);
					
			driver.findElement(By.cssSelector(arrPolpr[7])).clear();
			driver.findElement(By.cssSelector(arrPolpr[7])).sendKeys(Maxlength);
				
			driver.findElement(By.cssSelector(arrPolpr[8])).clear();
			driver.findElement(By.cssSelector(arrPolpr[8])).sendKeys(size);
				
			if(ValidationRequired.equalsIgnoreCase("Yes"))
			{
				System.out.println(arrPolpr[9]);
				driver.findElement(By.cssSelector(arrPolpr[9])).click();
			}
			if(PostBack.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[10])).click();
			}
			if(ItemInfo.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[11])).click();
			}
			if(Visible.equalsIgnoreCase("No"))
			{
				driver.findElement(By.cssSelector(arrPolpr[12])).click();
			}
			if(ReadOnly.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[13])).click();
			}
		
			 count++;
		}		
	}
	
}
public void dragPasswordinCol(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	System.out.println("am here");
	if(dt.equalsIgnoreCase("Passwordlist"))
	{
		
		Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
		Sheet sheet1 = wb1.getSheet("Passwordlist");
		int rows1 = sheet1.getRows();
		
		int cols1 = sheet1.getColumns();
		System.out.println(cols1);
		
		System.out.println(rows1);
		String data="";
		String Label, WaterMarkText,Name,Value,Maxlength,size,ValidationRequired,PostBack,ItemInfo,Visible,ReadOnly="";
		int count = 1;
		for (int row1 = 1; row1 < rows1; row1++) 
		{
			
			Label = sheet1.getCell(0, row1).getContents();
			WaterMarkText = sheet1.getCell(1, row1).getContents();
			Name = sheet1.getCell(2, row1).getContents();
			Value = sheet1.getCell(3, row1).getContents();
			Maxlength = sheet1.getCell(4, row1).getContents();
			size = sheet1.getCell(5, row1).getContents();
			ValidationRequired = sheet1.getCell(6, row1).getContents();
			PostBack = sheet1.getCell(7, row1).getContents();
			ItemInfo = sheet1.getCell(8, row1).getContents();
			Visible = sheet1.getCell(9, row1).getContents();
			ReadOnly = sheet1.getCell(10, row1).getContents();
			
			System.out.println(arrPolpr[0]);
			
			System.out.println(driver.findElement(By.xpath(arrPolpr[0])).isDisplayed());
			WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
			WebElement we1 = driver.findElement(By.cssSelector(arrPolpr[1]));
			Actions act = new Actions(driver);
			act.clickAndHold(we).build().perform();
			act.release(we1).build().perform();
			Thread.sleep(5000);
			
			driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div[2]/div/div/div/div/div/div[2]/div/div/div[1]/div/div["+count+"]")).click();
					  
			driver.findElement(By.cssSelector(arrPolpr[3])).clear();
			driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(Label);
				
			driver.findElement(By.cssSelector(arrPolpr[4])).clear();
			driver.findElement(By.cssSelector(arrPolpr[4])).sendKeys(WaterMarkText);
					
			driver.findElement(By.cssSelector(arrPolpr[5])).clear();
			driver.findElement(By.cssSelector(arrPolpr[5])).sendKeys(Name);
					
			driver.findElement(By.cssSelector(arrPolpr[6])).clear();
			driver.findElement(By.cssSelector(arrPolpr[6])).sendKeys(Value);
					
			driver.findElement(By.cssSelector(arrPolpr[7])).clear();
			driver.findElement(By.cssSelector(arrPolpr[7])).sendKeys(Maxlength);
				
			driver.findElement(By.cssSelector(arrPolpr[8])).clear();
			driver.findElement(By.cssSelector(arrPolpr[8])).sendKeys(size);
				
			if(ValidationRequired.equalsIgnoreCase("Yes"))
			{
				//System.out.println(arrPolpr[9]);
				driver.findElement(By.cssSelector(arrPolpr[9])).click();
			}
			if(PostBack.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[10])).click();
			}
			if(ItemInfo.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[11])).click();
			}
			if(Visible.equalsIgnoreCase("No"))
			{
				driver.findElement(By.cssSelector(arrPolpr[12])).click();
			}
			if(ReadOnly.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[13])).click();
			}
		
			 count++;
		}		
	}
	
}

public void verifyTextBoxes(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split(DELIM);
	int flag=0;
	String str="";
	String data="";
	if(dt.equalsIgnoreCase("TextBoxes"))
	{
		Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
		Sheet sheet1 = wb1.getSheet("TextBoxes");
		int rows1 = sheet1.getRows();
		
		
		int count = 1;
		for (int row1 = 1; row1 < rows1; row1++) 
		{
			data = sheet1.getCell(0, row1).getContents();
			str = driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div[2]/div/div/div/div/div/div[2]/div["+count+"]")).getText();
			if(str.equals(data))
			{
				System.out.println("pass");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Label Text ",data,data+ " - Label Text is matching",str+ " - Label Text is matching", "Pass");
				
			}
			else
			{
				System.out.println("fail");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Label Text ",data,data+ " - Label Text is matching",str+ " - Label Text is not matching", "Fail");
				
			}
			count++;
		}
			
	}
	
	
}



public void dragPassword(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	//String[] arrPoldt = dt.split("%");
	
	if(dt.equalsIgnoreCase("Passwordlist"))
	{
		
		Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
		Sheet sheet1 = wb1.getSheet("Passwordlist");
		int rows1 = sheet1.getRows();
		
		String data="";
		String Label, WaterMarkText,Name,Value,Maxlength,size,ValidationRequired,PostBack,ItemInfo,Visible,ReadOnly="";
		int count = 1;
		for (int row1 = 1; row1 < rows1; row1++) 
		{
			
			Label = sheet1.getCell(0, row1).getContents();
			WaterMarkText = sheet1.getCell(1, row1).getContents();
			Name = sheet1.getCell(2, row1).getContents();
			Value = sheet1.getCell(3, row1).getContents();
			Maxlength = sheet1.getCell(4, row1).getContents();
			size = sheet1.getCell(5, row1).getContents();
			ValidationRequired = sheet1.getCell(6, row1).getContents();
			PostBack = sheet1.getCell(7, row1).getContents();
			ItemInfo = sheet1.getCell(8, row1).getContents();
			Visible = sheet1.getCell(9, row1).getContents();
			ReadOnly = sheet1.getCell(10, row1).getContents();
			
			WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
			WebElement we1 = driver.findElement(By.cssSelector(arrPolpr[1]));
			Actions act = new Actions(driver);
			act.clickAndHold(we).build().perform();
			act.release(we1).build().perform();
			Thread.sleep(5000);
			//System.out.println(driver.findElement(By.cssSelector("div.Password[ng-class='Item.Type']")).isDisplayed());
			//driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div[2]/div/div/div/div/div/div[2]/div["+count+"]")).click();
			List<WebElement> wee = driver.findElements(By.cssSelector("div.Password[ng-class='Item.Type']"));//.subList(row1, rows1);
			int len=wee.size();
			System.out.println(len);
			for(int j=1;j<len;j++)
			{
				
				System.out.println(wee.get(j).findElement(By.cssSelector("div.Password[ng-class='Item.Type']")).isDisplayed());
				
			driver.findElement(By.cssSelector("div.Password[ng-class='Item.Type']")).click();
			//System.out.println(wee.size());
			
					  
			driver.findElement(By.cssSelector(arrPolpr[3])).clear();
			driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(Label);
				
			driver.findElement(By.cssSelector(arrPolpr[4])).clear();
			driver.findElement(By.cssSelector(arrPolpr[4])).sendKeys(WaterMarkText);
					
			driver.findElement(By.cssSelector(arrPolpr[5])).clear();
			driver.findElement(By.cssSelector(arrPolpr[5])).sendKeys(Name);
					
			driver.findElement(By.cssSelector(arrPolpr[6])).clear();
			driver.findElement(By.cssSelector(arrPolpr[6])).sendKeys(Value);
					
			driver.findElement(By.cssSelector(arrPolpr[7])).clear();
			driver.findElement(By.cssSelector(arrPolpr[7])).sendKeys(Maxlength);
				
			driver.findElement(By.cssSelector(arrPolpr[8])).clear();
			driver.findElement(By.cssSelector(arrPolpr[8])).sendKeys(size);
				
			if(ValidationRequired.equalsIgnoreCase("Yes"))
			{
				//System.out.println(arrPolpr[9]);
				driver.findElement(By.cssSelector(arrPolpr[9])).click();
			}
			if(PostBack.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[10])).click();
			}
			if(ItemInfo.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[11])).click();
			}
			if(Visible.equalsIgnoreCase("No"))
			{
				driver.findElement(By.cssSelector(arrPolpr[12])).click();
			}
			if(ReadOnly.equalsIgnoreCase("Yes"))
			{
				driver.findElement(By.cssSelector(arrPolpr[13])).click();
			}
		
			 count++;
		}	}
		
	}
	
}

public void verifyPassword(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	//System.out.println("am here");
	
	String[] arrPolpr = pr.split(DELIM);
	//String[] arrPoldt = dt.split("%");
	
	if(dt.equalsIgnoreCase("Passwordlist"))
	{
		//System.out.println("am here");
		Workbook wb1 = Workbook.getWorkbook(new File(DRIVEREXCEL));
		Sheet sheet1 = wb1.getSheet("Passwordlist");
		int rows1 = sheet1.getRows();
		
		
		//System.out.println(cols1);
		
		//System.out.println(rows1);
		String data="";
		int count = 1;
		for (int row1 = 1; row1 < rows1; row1++) 
		{
			
			
			data = sheet1.getCell(0, row1).getContents();
			//System.out.println(data);
			String str = driver.findElement(By.xpath("/html/body/form/div/div[1]/div/div[2]/div/div/div/div/div/div[2]/div["+count+"]")).getText();
			//System.out.println(str);
			if(str.equalsIgnoreCase(data))
			{
				System.out.println("pass");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Label Text ",data,data+ " - Label Text is matching",str+ " - Label Text is matching", "Pass");
			}
			else
			{
				System.out.println("fail");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Label Text ",data,data+ " - Label Text is matching",str+ " - Label Text is not matching", "Fail");
			}
			count++;
		}
			
	}
	
	
}
public void dragTextarea(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	
	
	String[] arrPoldt = dt.split("%");
	
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.cssSelector(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
	
	driver.findElement(By.cssSelector(arrPolpr[2])).click();
	
	
	driver.findElement(By.cssSelector(arrPolpr[3])).clear();
	driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(arrPoldt[0]);
	
	
	driver.findElement(By.cssSelector(arrPolpr[4])).clear();
	//driver.findElement(By.cssSelector(arrPolpr[4])).sendKeys(dt);
	
	driver.findElement(By.cssSelector(arrPolpr[5])).clear();
	driver.findElement(By.cssSelector(arrPolpr[5])).sendKeys(arrPoldt[1]);
	
	driver.findElement(By.cssSelector(arrPolpr[7])).click();
	
	
}





public void dragRadioGroup(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	String[] arrPolpr = pr.split("%");
	
	
	String[] arrPoldt = dt.split("%");
	
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.cssSelector(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
	
	driver.findElement(By.cssSelector(arrPolpr[2])).click();
	
	
	driver.findElement(By.cssSelector(arrPolpr[3])).clear();
	driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(arrPoldt[1]);
	
	
	//driver.findElement(By.cssSelector(arrPolpr[4])).clear();
	//driver.findElement(By.cssSelector(arrPolpr[4])).sendKeys(dt);
	
	driver.findElement(By.cssSelector(arrPolpr[5])).clear();
	driver.findElement(By.cssSelector(arrPolpr[5])).sendKeys(arrPoldt[2]);
	
	driver.findElement(By.cssSelector(arrPolpr[7])).click();
	
	
	//System.out.println(arrPolpr[8]);
	//System.out.println(driver.findElement(By.cssSelector(arrPolpr[8])).isDisplayed());
	
	System.out.println(driver.findElement(By.cssSelector(arrPolpr[8])).isDisplayed());
	
	List<WebElement> wee1 = driver.findElements(By.cssSelector(arrPolpr[8]));
	System.out.println(wee1.size());
	
	Integer totopt = Integer.parseInt(arrPoldt[0]);
	System.out.println(totopt);
	
	int dtv=3;
	int dev = totopt-wee1.size();
	
	if(totopt>wee1.size())
	{
		
		for (int k=0;k<dev;k++)
		{
			
			driver.findElement(By.xpath(arrPolpr[11])).click();
		}
		
		for(int i=2;i<totopt;i++)
		{
			
			List<WebElement> wee = driver.findElements(By.xpath("/html/body/form/div/div[3]/div/div/label['"+i+"']"));
			//System.out.println(wee.size());
			int count=1; 									 //for 4 elements, select option, option name, +, -
			for(int j=1;j<wee.size();j++)
			{
				
				
					
				
					//driver.findElement(By.cssSelector(arrPolpr[9])).click();
					wee.get(count).click();
					Actions act1 = new Actions(driver);
					act1.sendKeys(Keys.TAB).build().perform();
					//Thread.sleep(5000);
				
					count++;
				
				
				
					//System.out.println(arrPoldt[dtv]);
					driver.findElement(By.cssSelector(arrPolpr[9])).clear();
					act1.sendKeys(arrPoldt[dtv]);
				
					//System.out.println(count);
				
				
					act1.sendKeys(Keys.TAB).build().perform();
					act1.sendKeys(Keys.TAB).build().perform();
					dtv++;
				}
				
			}

		
	
	}
	
	
	if(totopt==wee1.size())
	{
		for(int i=2;i<totopt;i++)
		{
			List<WebElement> wee = driver.findElements(By.xpath("/html/body/form/div/div[3]/div/div/label['"+i+"']"));
			System.out.println(wee.size());
			int count=1; 									 //for 4 elements, select option, option name, +, -
			for(int j=1;j<wee.size();j++)
			{
				
				//driver.findElement(By.cssSelector(arrPolpr[9])).click();
				wee.get(count).click();
				Actions act1 = new Actions(driver);
				act1.sendKeys(Keys.TAB).build().perform();
				//Thread.sleep(5000);
				
				count++;
				
				
				
				//System.out.println(arrPoldt[dtv]);
				driver.findElement(By.cssSelector(arrPolpr[9])).clear();
				act1.sendKeys(arrPoldt[dtv]);
				
				//System.out.println(count);
				
				
				act1.sendKeys(Keys.TAB).build().perform();
				act1.sendKeys(Keys.TAB).build().perform();
				dtv++;
			}

		}
	}
	
	if(totopt<wee1.size())
	{
		
		for(int i=2;i<=totopt;i++)
		{
			driver.findElement(By.xpath(arrPolpr[10])).click();
			List<WebElement> wee = driver.findElements(By.xpath("/html/body/form/div/div[3]/div/div/label['"+i+"']"));
			//System.out.println(wee.size());
			int count=1; 									 //for 4 elements, select option, option name, +, -
			for(int j=1;j<wee.size();j++)
			{
				wee.get(count).click();
				Actions act1 = new Actions(driver);
				act1.sendKeys(Keys.TAB).build().perform();
				count++;
				System.out.println(arrPoldt[dtv]);
				driver.findElement(By.cssSelector(arrPolpr[9])).clear();
				act1.sendKeys(arrPoldt[dtv]);
				System.out.println(count);
				act1.sendKeys(Keys.TAB).build().perform();
				act1.sendKeys(Keys.TAB).build().perform();
				dtv++;
			}

		
		}
		
		
	}

}


public void dragRow(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	Thread.sleep(5000);
	String[] arrPolpr = pr.split("%");
	driver.switchTo().frame(0);
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}
public void dragRowinForm(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	
	//System.out.println("am here");
	Thread.sleep(5000);
	String[] arrPolpr = pr.split("%");
	//driver.switchTo().frame(0);
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}
public void dragMasterForm(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
	driver.findElement(By.xpath(arrPolpr[2])).click();
	driver.findElement(By.cssSelector(arrPolpr[3])).clear();
	driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(dt);
	driver.findElement(By.cssSelector(arrPolpr[4])).click();
	driver.findElement(By.xpath(arrPolpr[2])).click();

}

public void dragForm(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
	driver.findElement(By.xpath(arrPolpr[2])).click();
	driver.findElement(By.cssSelector(arrPolpr[3])).clear();
	driver.findElement(By.cssSelector(arrPolpr[3])).sendKeys(dt);
	driver.findElement(By.xpath(arrPolpr[2])).click();

}

public void dragTab(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}


public void dragSection(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}

public void dragLineBreak(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}

public void dragGroupPanel(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}

public void dragParagraph(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);
}

public void dragColoumn(String pr, String dt)throws InterruptedException, IOException, BiffException   {
	String[] arrPolpr = pr.split("%");
	WebElement we = driver.findElement(By.xpath(arrPolpr[0]));
	WebElement we1 = driver.findElement(By.xpath(arrPolpr[1]));
	Actions act = new Actions(driver);
	act.clickAndHold(we).build().perform();
	act.release(we1).build().perform();
	Thread.sleep(5000);

}






	public void verifyLayouts(String pr, String dt)throws InterruptedException, IOException, BiffException   {
		
		
		
		
		String[] arrPolpr = pr.split(DELIM);
		
		
		WebElement we = driver.findElement(By.className(arrPolpr[0]));
		
		List<WebElement> wetr = we.findElements(By.tagName("tr")) ;
		int flag=0;
		for(int i=1;i<wetr.size();i++)
		{
			
			WebElement wetd = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['"+i+"'+1]/td[1]/a"));
			
			String str = wetd.getText();
				
			if(str.equals(dt))
				
			{
				flag++;
				System.out.println("pass");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",dt,dt+ " - Master Data is available",str+ " - Master Data is available", "Pass");
				
				
			}
			
				
		}
		
		if(flag==0)
		{
			
			System.out.println("Fail");
			writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Master Data",dt,dt+ " - Master Data is available",dt+ " - Master Data is not available", "Fail");
			
			
		}
		
	}

	public void clickByLink(String pr, String dt)throws InterruptedException, IOException   {
		
		//Thread.sleep(10000);
		WebElement we = driver.findElement(By.linkText(pr));
		we.click();
		
		
		
	}
	
	public void verifyTaskfromGroupTask(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		Actions act = new Actions(driver);
		Thread.sleep(5000);
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[4]));
		
		for(int k=0;k<we.size();k++)
		{
			
			we.get(k).findElement(By.cssSelector(ArrPolpr[5])).sendKeys(dt);
				
			WebElement webody = we.get(k).findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",str+ " - Task is available under My Group Task section", "Pass");
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr['"+i+"']/td[3]")).getText();
					System.out.println(str);
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",str+ " - Task is available under My Group Task section", "Pass");
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",dt+ " - Task is not available under My Group Task section", "Fail");
			
			}
			
			
			
		}
	}
	
	
	
	
	
	
	
	public void checkJobinRepository(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		ArrPoldt = dt.split("%");
		
		driver.findElement(By.cssSelector(ArrPolpr[4])).click();
		
		Thread.sleep(5000);
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[5]));
		
		for(int k =0;k<we.size();k++)
		{
			
			//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			we.get(k).findElement(By.cssSelector(ArrPolpr[6])).clear();
			
			we.get(k).findElement(By.cssSelector(ArrPolpr[6])).sendKeys(dt);
			
			WebElement webody = we.get(k).findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under Repository",dt,dt+ " - Task is available under Repository",str+ " - Task is available under Repository", "Pass");
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int j =0;j<wetr.size();j++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]")).getText();
					
					System.out.println(str);
					
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under Repository",dt,dt+ " - Task is available under Repository",str+ " - Task is available under Repository", "Pass");
						break;
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under Repository",dt,dt+ " - Task is available under Repository",dt+ " - Task is not available under Repository", "Fail");
				//break;
			}
			
		//	break;
			
		}
		
		
}
	
	
	
	
	
	public void enterAddTimes(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		ArrPoldt = dt.split("%");
		
		driver.findElement(By.cssSelector(ArrPolpr[0])).sendKeys(ArrPoldt[0]);
		driver.findElement(By.cssSelector(ArrPolpr[1])).sendKeys(ArrPoldt[1]);
		driver.findElement(By.cssSelector(ArrPolpr[2])).sendKeys(ArrPoldt[2]);
		driver.findElement(By.cssSelector(ArrPolpr[3])).click();
	}
	
	
	
	
	public void clickDownload(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		ArrPoldt = dt.split("%");
		
		//driver.findElement(By.cssSelector(ArrPolpr[4])).click();
		
		Thread.sleep(5000);
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[5]));
		
		for(int k =0;k<we.size();k++)
		{
			
			//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			we.get(k).findElement(By.cssSelector(ArrPolpr[6])).clear();
			
			we.get(k).findElement(By.cssSelector(ArrPolpr[6])).sendKeys(dt);
			
			WebElement webody = we.get(k).findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under Repository",dt,dt+ " - Task is available under Repository",str+ " - Task is available under Repository", "Pass");
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int j =0;j<wetr.size();j++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]")).getText();
					
					System.out.println(str);
					
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under Repository",dt,dt+ " - Task is available under Repository",str+ " - Task is available under Repository", "Pass");
						break;
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under Repository",dt,dt+ " - Task is available under Repository",dt+ " - Task is not available under Repository", "Fail");
				//break;
			}
			
		//	break;
			
		}
		
		
}
	
	
	
	public void clickStartNewComment(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(2000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			driver.findElement(By.cssSelector(ArrPolpr[0])).click();
			WebElement we = driver.findElement(By.cssSelector("div._modal"));
			Thread.sleep(2000);
			//System.out.println(we.findElement(By.cssSelector("span.selected-list")).isDisplayed());
			
			we.findElement(By.cssSelector("div.custom-select")).click();
			
			Thread.sleep(5000);
			WebElement welst = we.findElement(By.cssSelector("div.custom-select"));
			
			WebElement wedls = welst.findElement(By.cssSelector("div.select-list-container"));
			List<WebElement> uls = wedls.findElements(By.tagName("ul"));
			
			for(int i=0;i<uls.size();i++)
			{
				List<WebElement> lis = uls.get(i).findElements(By.tagName("li"));
				
				for(int j=0;j<lis.size();j++)
				{
					System.out.println(lis.get(j).getText());
					System.out.println(lis.get(j).getText().equalsIgnoreCase(" Query"));
					if(lis.get(j).getText().equalsIgnoreCase(" Query"))
					{
						Actions act = new Actions(driver);
						act.moveToElement(lis.get(j)).build().perform();
						act.click(lis.get(j)).build().perform();
						
						System.out.println(we.findElement(By.cssSelector("div.form-group[ng-if='IsSubCategory.Properties.Visible']")).isDisplayed());
						WebElement wescat = we.findElement(By.cssSelector("div.form-group[ng-if='IsSubCategory.Properties.Visible']"));
						
						wescat.findElement(By.cssSelector("div.custom-select")).click();
						
						WebElement welst1 = wescat.findElement(By.cssSelector("div.custom-select"));
						
						WebElement wedls1 = welst1.findElement(By.cssSelector("div.select-list-container"));
						List<WebElement> uls1 = wedls1.findElements(By.tagName("ul"));
						
						for(int ii=0;ii<uls1.size();ii++)
						{
							List<WebElement> lis1 = uls1.get(ii).findElements(By.tagName("li"));
							
							for(int jj=0;jj<lis1.size();jj++)
							{
								System.out.println(lis1.get(jj).getText());
								System.out.println(lis1.get(jj).getText().equalsIgnoreCase(" Technical Query"));
								if(lis1.get(jj).getText().equalsIgnoreCase(" Technical Query"))
								{
									Actions act1 = new Actions(driver);
									act.moveToElement(lis1.get(jj)).build().perform();
									act.click(lis1.get(jj)).build().perform();
									
									driver.findElement(By.cssSelector("input.form-control[name='txtCommentSubject']")).sendKeys("Comment Subject");
									driver.findElement(By.cssSelector("input.form-control[name='txtComments']")).sendKeys("Comments texts");
									driver.findElement(By.cssSelector("input.btn[value='Submit']")).click();
								}
							}break;
						}
						
					}
					
				}
				break;	
				
				
			}
	}
	
	public void clickSupportingDoc(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(10000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			driver.switchTo().frame(0);
			WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
			System.out.println("am here");
			List<WebElement> weul = we.findElements(By.tagName("ul"));
			for(WebElement weulall:weul)
			{
				List<WebElement> weli = weulall.findElements(By.tagName("li"));
				for(WebElement weliall:weli)
				{
					System.out.println((weliall.getText()));
					if(weliall.getText().equalsIgnoreCase("Supporting Documents"))
					{
						weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
						Thread.sleep(5000);
						//*************************for submit un comment this code
						//driver.findElement(By.xpath(ArrPolpr[2])).click();
						break;
					}
							
				}break;	
			}
		
		
		
		
	}

	
	
	
	public void clickSupportingDocUpload(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(3000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			
			driver.findElement(By.cssSelector("input.btn[value='Upload']")).click();
			
			driver.switchTo().frame(0);
			
			
			WebElement we21 = driver.findElement(By.cssSelector("input.file[name='file']"));
			
			we21.click();
		
			Runtime.getRuntime().exec("D:\\Users\\kartheeswaran.r\\Desktop\\Up.exe");
			
			
			
			
		
			
			Thread.sleep(10000);
	
	}
	public void clickSupportingDocDownload(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(3000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			
			System.out.println("am here");
			
			
			
			driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).click();
			
			
			driver.findElement(By.cssSelector("input.btn[value='Download']")).click();
			
			//driver.switchTo().frame(0);
			Thread.sleep(5000);
			Actions act = new Actions(driver);
			
			act.sendKeys(Keys.ENTER).build().perform();
			
			
			//act.keyDown(Keys.ENTER).build().perform();
			//act.keyUp(Keys.ENTER).build().perform();
			
			
			Thread.sleep(10000);
	
	}
	
	public void clickSupportingSendMail(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(3000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			
			System.out.println("am here");
			
			
			
			driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).click();
			
			
			driver.findElement(By.cssSelector("input.btn[value='Send Mail']")).click();
			
			//driver.switchTo().frame(0);
			Thread.sleep(5000);
			Actions act = new Actions(driver);
			
			act.sendKeys(Keys.ENTER).build().perform();
			
			
			//act.keyDown(Keys.ENTER).build().perform();
			//act.keyUp(Keys.ENTER).build().perform();
			
			
			Thread.sleep(10000);
	
	}
	public void clickComments(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(10000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			driver.switchTo().frame(0);
			WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
			List<WebElement> weul = we.findElements(By.tagName("ul"));
			for(WebElement weulall:weul)
			{
				List<WebElement> weli = weulall.findElements(By.tagName("li"));
				for(WebElement weliall:weli)
				{
					if(weliall.getText().equalsIgnoreCase("Comments"))
					{
						weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
						Thread.sleep(5000);
						//*************************for submit un comment this code
						//driver.findElement(By.xpath(ArrPolpr[2])).click();
						break;
					}
						
				}break;	
			}
		
	}
	
	
	
	
	public void clickJobTaskMap(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
			Thread.sleep(10000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			driver.switchTo().frame(0);
			WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
			List<WebElement> weul = we.findElements(By.tagName("ul"));
			for(WebElement weulall:weul)
			{
				List<WebElement> weli = weulall.findElements(By.tagName("li"));
				for(WebElement weliall:weli)
				{
					if(weliall.getText().equalsIgnoreCase("Job Task Map"))
					{
						weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
						Thread.sleep(5000);
						//*************************for submit un comment this code
						//driver.findElement(By.xpath(ArrPolpr[2])).click();
						break;
					}
						
				}break;	
			}
	}
	
	
	
	
	public void clickProofInfo(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
			Thread.sleep(10000);
			String[] ArrPolpr,ArrPoldt;
			ArrPolpr = pr.split("%");
			//driver.switchTo().frame(0);
			WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
			List<WebElement> weul = we.findElements(By.tagName("ul"));
			for(WebElement weulall:weul)
			{
				List<WebElement> weli = weulall.findElements(By.tagName("li"));
				for(WebElement weliall:weli)
				{
					if(weliall.getText().equalsIgnoreCase("Proof Info"))
					{
						weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
						Thread.sleep(5000);
						//*************************for submit un comment this code
						//driver.findElement(By.xpath(ArrPolpr[2])).click();
						break;
					}
						
				}break;	
			}
	}
	
	public void verifyTaskfromMyTask(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		
		Thread.sleep(5000);
		//ArrPoldt = dt.split("%");
		//System.out.println(ArrPolpr[4]);
		//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[4])).isDisplayed());
		
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[4]));
		
		for(WebElement wee : we)
		{
			
			//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			wee.findElement(By.cssSelector(ArrPolpr[5])).clear();
			
			wee.findElement(By.cssSelector(ArrPolpr[5])).sendKeys(dt);
			
			WebElement webody = wee.findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Task section",dt,dt+ " - Task is available under My Task section",str+ " - Task is available under My Task section", "Pass");
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int j =0;j<wetr.size();j++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]")).getText();
					
					System.out.println(str);
					
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Task section",dt,dt+ " - Task is available under My Task section",str+ " - Task is available under My Task section", "Pass");
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Task section",dt,dt+ " - Task is available under My Task section",dt+ " - Task is not available under My Task section", "Fail");
			
			}
			
			
			
		}
}
	
	

	public void verifyTask(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		Thread.sleep(10000);
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		//System.out.println(ArrPolpr[4]);
		Thread.sleep(5000);
		System.out.println(driver.findElement(By.cssSelector(ArrPolpr[4])).isDisplayed());
		
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[4]));
		
		for(WebElement wee : we)
		{
			
			System.out.println(driver.findElement(By.xpath(ArrPolpr[5])).isDisplayed());
			
			wee.findElement(By.xpath(ArrPolpr[5])).clear();
			
			wee.findElement(By.xpath(ArrPolpr[5])).sendKeys(dt);
			
			WebElement webody = wee.findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",str+ " - Task is available under My Group Task section", "Pass");
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr['"+i+"']/td[3]")).getText();
					System.out.println(str);
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",str+ " - Task is available under My Group Task section", "Pass");
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",dt+ " - Task is not available under My Group Task section", "Fail");
			
			}
			
		}
			

	}
	
	
	
	public void assignToMeFunc(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		Thread.sleep(10000);
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		//System.out.println(ArrPolpr[4]);
		//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[4])).isDisplayed());
		
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[4]));
		
		for(WebElement wee : we)
		{
			
			//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			wee.findElement(By.cssSelector(ArrPolpr[5])).clear();
			
			wee.findElement(By.cssSelector(ArrPolpr[5])).sendKeys(dt);
			
			WebElement webody = wee.findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",str+ " - Task is available under My Group Task section", "Pass");
						
						//System.out.println(driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[17]/span[2]/span/a")).isDisplayed());
						
						
						// Make sure Reassign and AssignToMe both buttons are there or not, as per that we have to put this code//
						
						//System.out.println("am here");
						driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[17]/span[2]/span/a")).click();
						//System.out.println("am here");
						
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr['"+i+"']/td[3]")).getText();
					System.out.println(str);
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",str+ " - Task is available under My Group Task section", "Pass");
						
						System.out.println(driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr['"+i+"']/td[10]/span/span/a")).isDisplayed());
						
						driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr['"+i+"']/td[10]/span/span/a")).click();
						
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Group Task section",dt,dt+ " - Task is available under My Group Task section",dt+ " - Task is not available under My Group Task section", "Fail");
			
			}
			
		}
			

	}
	
	
	public void reassignToMeFunc(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		//ArrPoldt = dt.split("%");
		/*System.out.println(ArrPolpr[4]);
		System.out.println(driver.findElement(By.cssSelector(ArrPolpr[4])).isDisplayed());
		*/
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[4]));
		
		for(WebElement wee : we)
		{
			
			//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			wee.findElement(By.cssSelector(ArrPolpr[5])).sendKeys(dt);
			
			WebElement webody = wee.findElement(By.tagName("tbody"));
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i =0;i<wetr.size();i++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).getText();
					
				
					
					System.out.println(str);
					
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Task section",dt,dt+ " - Task is available under My Task section",str+ " - Task is available under My Task section", "Pass");
						
						
						System.out.println(driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[17]/span/span/a")).isDisplayed());
						
						driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[17]/span/span/a")).click();
						
						WebElement weusr = driver.findElement(By.cssSelector("select[name='_inp_0']"));
						
						Actions act = new Actions(driver);
						
						act.click(weusr).build().perform();
						
						Select sel = new Select(weusr);
						
						sel.selectByVisibleText("G4Admin");
						
						
						
						act.sendKeys(Keys.ENTER).build().perform();
						
						driver.findElement(By.cssSelector("input[value='Ok']")).click();
						
						
						

						
						break;
					}
				}
				break;
			}
			if(wetr.size()>1)
			{
				for(int j =0;j<wetr.size();j++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr['"+j+"']/td[3]")).getText();
					System.out.println(str);
					System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println("Pass");
						writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Task section",dt,dt+ " - Task is available under My Task section",str+ " - Task is available under My Task section", "Pass");
						driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[17]/span/span/a")).click();
						WebElement weusr = driver.findElement(By.cssSelector("select[name='_inp_0']"));
						
						Actions act = new Actions(driver);
						
						act.click(weusr).build().perform();
						
						Select sel = new Select(weusr);
						
						sel.selectByVisibleText("G4Admin");
						
						
						
						act.sendKeys(Keys.ENTER).build().perform();
						
						driver.findElement(By.cssSelector("input[value='Ok']")).click();
					}
				}
				break;
			}
			
			else
			{
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the following Task is available or not under My Task section",dt,dt+ " - Task is available under My Task section",dt+ " - Task is not available under My Task section", "Fail");
			
			}
			
			
			
		}
}
	
	
	public void clciktheTask(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		//Thread.sleep(10000);
		
		System.out.println("am here1");
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		//System.out.println(ArrPolpr[4]);
		//System.out.println(driver.findElement(By.cssSelector(ArrPolpr[4])).isDisplayed());
		
		List<WebElement> we = driver.findElements(By.cssSelector(ArrPolpr[4]));
		
		for(WebElement wee : we)
		{
			System.out.println("am here2");
			
//			System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
//			
//			System.out.println(driver.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			System.out.println(wee.findElement(By.cssSelector(ArrPolpr[5])).isDisplayed());
			
			wee.findElement(By.cssSelector(ArrPolpr[5])).clear();
		
			wee.findElement(By.cssSelector(ArrPolpr[5])).sendKeys(dt);
			
			//System.out.println(wee.findElement(By.tagName("tbody")).isDisplayed());
			
			System.out.println("am here1");
			
			WebElement webody = wee.findElement(By.tagName("tbody"));
			
			
			
			List<WebElement> wetr = webody.findElements(By.tagName("tr"));
			
			System.out.println(wetr.size());
			
			if(wetr.size()==1)
			{
				for(int i=0;i<wetr.size();i++)
				{
					
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[3]")).getText();
					//System.out.println(dt.equalsIgnoreCase(str));
					if(dt.equalsIgnoreCase(str))
					{
						System.out.println(driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[2]/a")).isDisplayed());
						driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr/td[2]/a")).click();
							
						break;
					}
				}
			}
				//break;
				
			if(wetr.size()>1)
			{
				for(int j =0;j<wetr.size();j++)
				{
					String str = driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]")).getText();
					if(dt.equalsIgnoreCase(str))
					{
						//System.out.println("am here");
							
						driver.findElement(By.xpath("/html/body/main/div/form/div/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/div[2]/table/tbody/tr[1]/td[2]/a")).click();
								
						break;
					}
						
						
					}
				}
			}
		
		
		

	
	}
	
	

	public void clickDraft(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		Thread.sleep(10000);
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		driver.switchTo().frame(0);
		WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
		List<WebElement> weul = we.findElements(By.tagName("ul"));
		for(WebElement weulall:weul)
		{
			List<WebElement> weli = weulall.findElements(By.tagName("li"));
			for(WebElement weliall:weli)
			{
				if(weliall.getText().equalsIgnoreCase("Draft"))
				{
					weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
					Thread.sleep(5000);
					//*************************for submit un comment this code
					//driver.findElement(By.xpath(ArrPolpr[2])).click();
					break;
				}
					break;	
			}break;	
		}
	
	}
	
	
	
	
	public void clickArtworkTab(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		Thread.sleep(10000);
		String[] ArrPolpr,ArrPoldt;
		ArrPolpr = pr.split("%");
		driver.switchTo().frame(0);
		WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
		List<WebElement> weul = we.findElements(By.tagName("ul"));
		for(WebElement weulall:weul)
		{
			List<WebElement> weli = weulall.findElements(By.tagName("li"));
			for(WebElement weliall:weli)
			{
				if(weliall.getText().equalsIgnoreCase("Artwork Upload"))
				{
					weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
					Thread.sleep(5000);
					WebElement we21 = driver.findElement(By.cssSelector(ArrPolpr[3]));
					we21.sendKeys("D:\\Users\\kartheeswaran.r\\Desktop\\pdf\\C0028.01.pdf");
					Actions act = new Actions(driver);
					act.sendKeys(Keys.TAB).build().perform();
					Thread.sleep(10000);
					
					//*************************for submit un comment this code
					driver.findElement(By.xpath(ArrPolpr[2])).click();
					
					break;
				}
					break;	
			}break;	
		}
	
	}
	
	
	
	public void clickArtworkReview(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		Thread.sleep(10000);
		String[] ArrPolpr,ArrPoldt;
		
		ArrPolpr = pr.split("%");
		
		ArrPoldt = dt.split("%");
		driver.switchTo().frame(0);
		WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
		List<WebElement> weul = we.findElements(By.tagName("ul"));
		for(WebElement weulall:weul)
		{
			List<WebElement> weli = weulall.findElements(By.tagName("li"));
			for(WebElement weliall:weli)
			{
				if(weliall.getText().equalsIgnoreCase("Artwork Review"))
				{
					weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
					Thread.sleep(5000);
					
					
					//driver.findElement(By.cssSelector(ArrPolpr[3])).sendKeys(ArrPoldt[0]);
					driver.findElement(By.cssSelector(ArrPolpr[4])).sendKeys(ArrPoldt[1]);
					
					
					
					
					System.out.println(driver.findElement(By.xpath(ArrPolpr[2])).isDisplayed());
					//*************************for submit un comment this code
					driver.findElement(By.xpath(ArrPolpr[2])).click();
					break;
				}
					break;	
			}break;	
		}
	
	}
	
public void clickArtworkApproval(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	
	{
		Thread.sleep(10000);
		String[] ArrPolpr,ArrPoldt;
		
		ArrPolpr = pr.split("%");
		
		ArrPoldt = dt.split("%");
		driver.switchTo().frame(0);
		WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
		List<WebElement> weul = we.findElements(By.tagName("ul"));
		for(WebElement weulall:weul)
		{
			List<WebElement> weli = weulall.findElements(By.tagName("li"));
			for(WebElement weliall:weli)
			{
				if(weliall.getText().equalsIgnoreCase("Artwork Approval"))
				{
					weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
					Thread.sleep(5000);
					
					
					//driver.findElement(By.cssSelector(ArrPolpr[3])).sendKeys(ArrPoldt[0]);
					driver.findElement(By.cssSelector(ArrPolpr[4])).sendKeys(ArrPoldt[1]);
					
					
					
					
					System.out.println(driver.findElement(By.xpath(ArrPolpr[2])).isDisplayed());
					//*************************for submit un comment this code
					driver.findElement(By.xpath(ArrPolpr[2])).click();
					break;
				}
					break;	
			}break;	
		}
	
	}

public void clickFinalApproval(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 

{
	Thread.sleep(10000);
	String[] ArrPolpr,ArrPoldt;
	
	ArrPolpr = pr.split("%");
	
	ArrPoldt = dt.split("%");
	driver.switchTo().frame(0);
	WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
	List<WebElement> weul = we.findElements(By.tagName("ul"));
	for(WebElement weulall:weul)
	{
		List<WebElement> weli = weulall.findElements(By.tagName("li"));
		for(WebElement weliall:weli)
		{
			if(weliall.getText().equalsIgnoreCase("Final Approval"))
			{
				weliall.findElement(By.cssSelector(ArrPolpr[1])).click();
				Thread.sleep(5000);
				
				
				//driver.findElement(By.cssSelector(ArrPolpr[3])).sendKeys(ArrPoldt[0]);
				driver.findElement(By.cssSelector(ArrPolpr[4])).sendKeys(ArrPoldt[1]);
				
				
				
				
				System.out.println(driver.findElement(By.xpath(ArrPolpr[2])).isDisplayed());
				//*************************for submit un comment this code
				driver.findElement(By.xpath(ArrPolpr[2])).click();
				break;
			}
				break;	
		}break;	
	}

}
	public void switchUserDropDwn(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] ArrPolpr;
		
		ArrPolpr = pr.split("%");
		
		Thread.sleep(5000);
	
		WebElement we = driver.findElement(By.cssSelector(ArrPolpr[0]));
		
		Actions act = new Actions(driver);
		
		act.click(we).build().perform();
		
		Select sel = new Select(we);
		
		sel.selectByVisibleText(dt);
		
		act.sendKeys(Keys.ENTER).build().perform();
		
	}
	
	

	public void verifySwitchedUser(String pr, String dt) throws InterruptedException, BiffException, IOException, AWTException 
	{
		
		String[] ArrPolpr;
		ArrPolpr = pr.split("%");
		Thread.sleep(5000);
		String str = driver.findElement(By.cssSelector(ArrPolpr[4])).getText();
		System.out.println(str);
		if(str.equalsIgnoreCase(dt))
		{
			System.out.println("Pass");
			writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the User has been switched or not",dt,dt+ " - User has been switched",str+ " - User has been switched", "Pass");
		}
		else
		{
			System.out.println("Fail");
			writeExcel.writeResult(RESULTEXCEL, ArrPolpr[2],ArrPolpr[3],"Verify the User has been switched or not",dt,dt+ " - User has been switched",dt+ " - User has not been switched", "Fail");
		}
	}
	
	
	public void findDescrinMytask(String pr, String dt)throws InterruptedException, IOException   {
	
		
	}
	
	
	/*
public void clickByLink1(String pr, String dt)throws InterruptedException, IOException   {
		
	
	String[] arrPolpr = pr.split("%");
	
	WebElement we3 = driver.findElement(By.className(arrPolpr[5]));
	WebElement we5=null;
	List<WebElement> we4 = we3.findElements(By.tagName("tr"));
	int len = we4.size();
	int flag = 5;
	for(int i = 1;i<len;i++)
	{	
		
		
		 we5 = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[1]"));
		String str = we5.getText();
		System.out.println(str);
			System.out.println(str.equalsIgnoreCase(dt));
			
			flag=0;
			if(str.equals(dt))
			{
				flag++;
				
				
		//System.out.println(driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[5]")).isEnabled());
		//System.out.println(driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[5]")).getText());
				
				WebElement we = driver.findElement(By.xpath("/html/body/main/div/table/tbody/tr['" + i + "'+1]/td[5]"));
				
				((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we);
				
				Actions act = new Actions(driver);
				
				
				Thread.sleep(2000);
				
				act.moveToElement(we).build().perform();
				act.click(we).build().perform();
				
				
				driver.findElement(By.xpath(arrPolpr[6])).click();
				
				Thread.sleep(2000);
				
				driver.findElement(By.xpath(arrPolpr[7])).click();
				
				
				
				Thread.sleep(5000);
				
				System.out.println("Pass");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Editted Master Data",dt,dt+ " - Master Data is editted Successfully",str+ " - Master Data is editted Successfully", "Pass");
				
				
				System.out.println(we5.getText());
				System.out.println(we5.isDisplayed());
				
				((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
				objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
															
				break;
			}
		}
			if(flag==0)
			{
				
				System.out.println("Fail");
				writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Editted Master Data",dt,dt+ " - Master Data is editted Successfully",dt+ " - Master Data is not editted Successfully", "Fail");
				
				((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView();", we5);
				objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
				
			}
			
		
	
		
		
		
		
	}
	*/
		
	/*
public void verifErroronLogon(String pr, String dt)throws InterruptedException, IOException   {
	Thread.sleep(10000);
	driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	
	String[] arrPolpr = pr.split(DELIM);
	String str = driver.findElement(By.cssSelector("div.alert-info[id='errr']")).getText();
	if (str.equals(dt)) {
		
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Error Text",dt,dt+ " text is displayed",str+ " - text is displayed", "Pass");
	objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	} else {
		
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[2],arrPolpr[3],"Verify Error Text",dt,dt+ " text is displayed",str+ " - text is displayed", "Fail");
	objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	//driver.close();
	}		
}	
	*/
public void mouseOver(String pr, String dt)throws InterruptedException, IOException   {
	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	Actions action = new Actions(driver);
	WebElement we = driver.findElement(By.xpath("//li[@id='wp-admin-bar-new-content']/a/span[2]"));
	action.moveToElement(we).moveToElement(driver.findElement(By.xpath("//a[contains(text(),'Achievement Type')])[2]"))).click().build().perform();

} 



public void mouseOverregister(String pr, String dt)throws InterruptedException, IOException   {
	driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	Actions action = new Actions(driver);
	WebElement we = driver.findElement(By.xpath("//a[contains(text(),'TestUser1')]"));
	action.moveToElement(we).moveToElement(driver.findElement(By.linkText("Activate"))).click().build().perform(); 

} 
public void selectbybadgeRole(String pr, String dt) throws InterruptedException, BiffException, IOException {
	
	Select oSelection = new Select(driver.findElement(By.id("role")));
	List<WebElement> oSize = oSelection.getOptions();
	int iListSize = oSize.size();
	
	
	driver.findElement(By.id(pr)).sendKeys(dt);
	WebElement dd= driver.findElement(By.id(pr));//pr
	Select mySelect= new Select(dd);
	mySelect.selectByVisibleText(dt);//dt

}
	
	//Reading frome Excel 
	/*
	String FilePath = "D:\\Autofiles\\Testdata.xls";
	 
	FileInputStream fs = new FileInputStream(FilePath);
	Workbook wb = Workbook.getWorkbook(fs);

	// TO get the access to the sheet
	Sheet sh = wb.getSheet("Sheet1");

	// To get the number of rows present in sheet
	int totalNoOfRows = sh.getRows();

	// To get the number of columns present in sheet
	int totalNoOfCols = sh.getColumns();

	for (int row = 0; row < totalNoOfRows; row++) { 

		for (int col = 0; col < totalNoOfCols; col++) {
			System.out.print(sh.getCell(col, row).getContents() + "\t");
			WebElement dd= driver.findElement(By.id(pr));//pr
			System.out.print(sh.getCell(col, row).getContents() + "\t");
			Select mySelect= new Select((WebElement) sh.getCell(col, row));
			
			
		}
		System.out.println();
	}
}
*/
		
	//driver.findElement(By.id(pr)).clear();
//driver.findElement(By.id(pr)).sendKeys(dt);
//WebElement dd= driver.findElement(By.id(pr));//pr
//Select mySelect= new Select(dd);
//mySelect.selectByVisibleText(dt);//dt






public void clickByclassname(String pr, String dt) throws InterruptedException { 
	
	//Thread.sleep(5000); 
	
	WebDriverWait ww = new WebDriverWait(driver, 10);
	ww.until(ExpectedConditions.presenceOfElementLocated(By.className(pr)));
		
	driver.findElement(By.className(pr)).click();
	
	//driver.findElement(By.xpath(pr)).click();
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	
}
public void refresh(String pr, String dt) throws InterruptedException { 
	
	driver.get(URL);
}

public void refreshcurrentpage(String pr, String dt) throws InterruptedException { 
	
	driver.navigate().refresh();
}
public void Delete(String pr, String dt) throws InterruptedException { 
	
	
	//User Delete
	
	driver.get("http://dev.asahitechnologies.com/badgeos/wp-admin/users.php");
	driver.findElement(By.id("user-search-input")).sendKeys("Student"); //To change the username
	driver.findElement(By.id("search-submit")).click();
	   
	   WebElement menu=driver.findElement(By.xpath("//a[contains(text(),'Student')]")); //Change
	   WebElement submenu=driver.findElement(By.className("submitdelete"));
	   Actions action = new Actions(driver);
	   action.moveToElement(menu).perform();
	   submenu.click();
	   driver.findElement(By.id("submit")).click();
	   
	   //User Delete
	      
	
	//Delete Achievement 
	   
	   WebDriverWait ww = new WebDriverWait(driver, 10);
		ww.until(ExpectedConditions.presenceOfElementLocated(By.linkText("BadgeOS")));
		
	driver.findElement(By.linkText("BadgeOS")).click();
	Actions builder = new Actions(driver); 
    WebElement mainmenu1 = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/ul/li[18]/a/div[3]"));
    builder.moveToElement(mainmenu1 ).build().perform();
    Thread.sleep(500); //add a wait
    WebElement submenu1=  driver.findElement(By.xpath("//a[contains(text(),'NewTest44')]"));
    builder.moveToElement(submenu1).click().build().perform(); 
    driver.findElement(By.xpath("//input[@id='cb-select-all-1']")).click();
    WebElement mySelectElm = driver.findElement(By.id("bulk-action-selector-top")); 
    Select mySelect= new Select(mySelectElm);
    mySelect.selectByVisibleText("Move to Trash");
    driver.findElement(By.id("doaction")).click();
    //Trash
    driver.findElement(By.cssSelector("li.trash > a")).click();
    driver.findElement(By.xpath("//input[@id='cb-select-all-1']")).click();
    WebElement mySelectDel = driver.findElement(By.id("bulk-action-selector-top")); 
    Select mySelectD= new Select(mySelectDel);
    mySelectD.selectByVisibleText("Delete Permanently");  
    driver.findElement(By.id("doaction")).click();
    
	
	//Delete Achievement Type
    WebElement mainmenu2 = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/ul/li[18]/a/div[3]"));
    builder.moveToElement(mainmenu2 ).build().perform();
    Thread.sleep(500); //add a wait
    WebElement submenu2=  driver.findElement(By.xpath("//a[contains(text(),'Achievement Types')]"));
    builder.moveToElement(submenu2).click().build().perform(); 
    driver.findElement(By.xpath("//input[@id='cb-select-all-1']")).click();
    WebElement mySelectDel1 = driver.findElement(By.id("bulk-action-selector-top")); 
    Select mySelectD1= new Select(mySelectDel1);
    mySelectD1.selectByVisibleText("Move to Trash");  
    driver.findElement(By.id("doaction")).click();
    driver.findElement(By.cssSelector("li.trash > a")).click();
    driver.findElement(By.xpath("//input[@id='cb-select-all-1']")).click();
    WebElement mySelectDel2 = driver.findElement(By.id("bulk-action-selector-top")); 
    Select mySelectD2= new Select(mySelectDel2);
    mySelectD2.selectByVisibleText("Delete Permanently");  
    driver.findElement(By.id("doaction")).click();
    // Result
    
    String[] arrPolpr = pr.split("%");
	  
  	System.out.println("Ach");
  	Thread.sleep(2000);
  	
  	 	
      String str = driver.findElement(By.xpath("//table/tbody/tr/td[1]")).getText();
      	
  		
  	if (str.equals(dt)) {
  		System.out.println("Ok score right");
  	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
  	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
  	} else {
  		System.out.println("Ok score wrong"); 
  	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
  	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
  	}
  	
  	
    
}

public void movetoElement(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000);  
	driver.navigate().to("http://dev.asahitechnologies.com/badgeos/wp-admin/edit.php?post_type=achievement-type");
	 
	/*WebDriverWait ww = new WebDriverWait(driver, 10);
	ww.until(ExpectedConditions.presenceOfElementLocated(By.linkText("BadgeOS")));
	
	driver.findElement(By.linkText("BadgeOS")).click();
	Actions builder = new Actions(driver); 
    WebElement mainmenu1 = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[2]/ul/li[18]/a/div[3]"));
    builder.moveToElement(mainmenu1 ).build().perform();
    Thread.sleep(500); //add a wait
    WebElement submenu1=  driver.findElement(By.xpath("//a[contains(text(),'Achievement Types')]"));
    builder.moveToElement(submenu1).click().build().perform();
	
    //driver.findElement(By.className(pr)).click();
	
	//driver.findElement(By.xpath(pr)).click();
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); */
	
}
public void movetoElementdropdown(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000);  
	 driver.findElement(By.xpath("//input[@value='Add New Step']")).click();
	 driver.findElement(By.xpath("//ul[@id='steps_list']/li/select/option[8]")).click();
	    Thread.sleep(1000);
	    driver.findElement(By.xpath("//select[4]/option[3]")).click();  
	    driver.findElement(By.xpath("//li/div[2]/input")).sendKeys("Complete the Steps");
	    driver.findElement(By.xpath("//input[@value='Save All Steps']")).click();
	
}

public void navigateto(String pr, String dt) throws InterruptedException { 
	
	//Thread.sleep(5000);  
	driver.navigate().to("http://dev.asahitechnologies.com/badgeos/wp-admin/users.php?page=bp-signups");
	
}

public void navigatetoUrl(String pr, String dt) throws InterruptedException { 
	
	//Thread.sleep(5000);  
	driver.navigate().to("http://dev.asahitechnologies.com/badgeos/"); 
	
} 
public void navigatetoprofile(String pr, String dt) throws InterruptedException { 
	
	//Thread.sleep(5000);  
	driver.navigate().to("http://dev.asahitechnologies.com/badgeos/members/testuser0001/profile/edit/group/1/"); 
	
}
public void verifyStudent(String pr, String dt) throws InterruptedException { 
	 
	//Thread.sleep(5000);  
	
	 String[] arrPolpr = pr.split("%"); 
	  
		
		Thread.sleep(2000); 
		 WebElement ele=driver.findElement(By.cssSelector("h1.entry-title"));  
		    String str = ele.getText().toString();
			System.out.println(str);  	
			
		if (str.equals(dt)) {
			System.out.println("Pass Student creation"); 
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("Fail Student creation"); 
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
	
	
	
	
}
public void verifyText(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000); 
	
	//WebDriverWait ww = new WebDriverWait(driver, 10);
	//ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
	List<WebElement> list = driver.findElements(By.xpath("//*[contains(text(),'" +"New user created" + "')]"));
	Assert.assertTrue("Text not found!", list.size() > 0);
	System.out.println("1. Admin User Created");	
	
	
	//driver.findElement(By.xpath(pr)).click();
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	
}
	
public void verifyTextS1(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000); 
	 
	//WebDriverWait ww = new WebDriverWait(driver, 10);
	//ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
	List<WebElement> list = driver.findElements(By.xpath("//*[contains(text(),'" +"S1" + "')]"));
	Assert.assertTrue("Text not found!", list.size() > 0);
	System.out.println("Text S1 Displayed");	 
	
	
	//driver.findElement(By.xpath(pr)).click();
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	
}
 

public void verifyTextipm(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000); 
	
	//WebDriverWait ww = new WebDriverWait(driver, 10);
	//ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
	
	 String[] arrPolpr = pr.split("%"); 
	  
		
		Thread.sleep(2000); 
		WebElement ele = driver.findElement(By.xpath("//header/h1")); 
		System.out.println(ele.getText()); 
		
	    String str = ele.getText().toString();
	    	
			 
		if (str.equals(dt)) {
			System.out.println("Pass IPM displayed");
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("Fail IPM displayed"); 
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
	
	
	//driver.findElement(By.xpath(pr)).click();
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	
}
	
public void clickonPublish(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000); 
	
	//WebDriverWait ww = new WebDriverWait(driver, 10);
	//ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
	 WebElement pub = driver.findElement(By.id("major-publishing-actions"));
     pub.findElement(By.xpath("//input[@id='publish']")).click(); 
	
	
	
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	
}

public void ParameterDelete(String pr, String dt) throws InterruptedException, BiffException, IOException {
	
	
		   
	   Workbook wb1 = Workbook.getWorkbook(new File(INPUTEXCEL));   
		Sheet sheet1 = wb1.getSheet("Users");
		//System.out.println(sheet1.g);
		int rows1 = sheet1.getRows();
		//int cols1 = sheet1.getColumns();
		
			for (int row1 = 0; row1 <rows1; row1++)  
			{
				//UserName
				System.out.println(sheet1.getCell(1, (row1+1)).getContents());
				String username = sheet1.getCell(1, (row1+1)).getContents();
				driver.findElement(By.id("user-search-input")).sendKeys(username); //To change the username
				driver.findElement(By.id("search-submit")).click();
				   
				    WebElement menu=driver.findElement(By.name("users[]")); 
				   //WebElement menu=driver.findElement(By.xpath("//a[contains(text(),username)]")); //Change 
				   WebElement submenu=driver.findElement(By.className("submitdelete"));
				   Actions action = new Actions(driver);
				   action.moveToElement(menu).perform();
				   submenu.click();
				   driver.findElement(By.id("submit")).click();
				   
			}
			driver.close();
	    
}

public void recentAchievement(String pr, String dt) throws InterruptedException { 
	
	List<WebElement> liElements = driver.findElements(By.xpath("//li[@id='toplevel_page_badgeos_badgeos']/a/div[3]"));

    System.out.println(liElements.size());

    for (int i = 1; i < liElements.size()+1; i++) { 
        WebElement linkElement = driver.findElement(By.linkText(dt));  
        linkElement.click();
     }
 
	  
		
}
 


public void achievementType(String pr, String dt) throws InterruptedException {
	
	Thread.sleep(3000); 
	
	 Select dropdown = new Select(driver.findElement(By.id("_badgeos_earned_by"))); 
    // dropdown.selectByVisibleText("Submission (Auto-accepted)");
     dropdown.selectByVisibleText(dt);	
	
	
	
}

public void ResultAchievementType(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(2000);  
	driver.navigate().to("http://dev.asahitechnologies.com/badgeos/wp-admin/edit.php?post_type=achievement-type"); 
	 
	/*
	
	List<WebElement> liElements = driver.findElements(By.xpath("//li[@id='toplevel_page_badgeos_badgeos']/a/div[3]"));
	
    System.out.println(liElements.size());

    for (int i = 1; i < liElements.size()+1; i++) {
        WebElement linkElement = driver.findElement(By.linkText("Achievement Types"));  
        linkElement.click();
     }
*/
    
    String[] arrPolpr = pr.split("%"); 
	  
	
	Thread.sleep(2000); 
	WebElement ele = driver.findElement(By.xpath("//table/tbody/tr/td[1]/strong/a")); 
	//ele.findElement(By.xpath("//a[contains(text(),'AT1')]")).getText(); 
	 	
    String str = ele.getText().toString();
    	
		
	if (str.equals(dt)) {
		System.out.println("Ok AT1");
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	} else {
		System.out.println("AT! Wrong"); 
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	}
	
	 
	
}

public void tSleep(String pr, String dt) throws InterruptedException { 
	
	Thread.sleep(3000); 
	
	
}
public void scrollUp(String pr, String dt) throws InterruptedException { 
	
	Thread.sleep(2000); 
	//	For Scroll up 
	
	JavascriptExecutor jse = (JavascriptExecutor)driver;
    jse.executeScript("scroll(0, -250);");
    
    //For save and publish 
    
   //// WebElement save = driver.findElement(By.id("minor-publishing-actions"));
   // WebElement save2= save.findElement(By.id("save-action"));
   // save2.findElement(By.id("save-post")).click();
   
    WebElement publish = driver.findElement(By.id("major-publishing-actions"));
    publish.findElement(By.xpath("//input[@id='publish']")).click(); 
	
	//Again clicking on plural value from menu
    
    List<WebElement> liElements1 = driver.findElements(By.xpath("//li[@id='toplevel_page_badgeos_badgeos']/a/div[3]")); 

    //System.out.println(liElements1.size()); 

    for (int i = 1; i < liElements1.size()+1; i++) { 
        WebElement linkElement1 = driver.findElement(By.linkText(dt)); //Change plural value 
        linkElement1.click();
     }

    System.out.println("2. Achievement created sucessfully");
	
}

public void scrollDown(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000);  
	
	JavascriptExecutor jse = (JavascriptExecutor)driver;
	jse.executeScript("window.scrollBy(0,250)", ""); 
	 
	
	
}
public void Verifyearnedachievement(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(5000); 
	
	//WebDriverWait ww = new WebDriverWait(driver, 10);
	//ww.until(ExpectedConditions.presenceOfElementLocated(By.xpath(pr)));
	List<WebElement> list = driver.findElements(By.xpath("//*[contains(text(),'" +"Thank you! Your submission has been received." + "')]"));
	Assert.assertTrue("Text not found!", list.size() > 0);
	System.out.println("3. You have earned this achievement!");	
	
	
	
}

public void scoreResult(String pr, String dt) throws InterruptedException {    
	
String[] arrPolpr = pr.split("%");  
 
//String str = driver.findElement(By.xpath("//*[contains(text(),'" +"New user created. Edit user" + "')]")).getText();

String str = driver.findElement(By.xpath("//*[contains(text(),'" +"You have earned this achievement!" + "')]")).getText(); 
//String str =driver.findElement(By.xpath("message/p")).getText();
System.out.println(str);
if (str.equals(dt)) {
	System.out.println("Ok score right");
writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
} else {
	System.out.println("Ok score wrong"); 
writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
}
}	
public void close(String pr, String dt) throws InterruptedException { 
	driver.close();
	
} 

public void submit(String pr, String dt) throws InterruptedException { 
	  WebElement iframeMsg = driver.findElement(By.xpath("//form/fieldset/p/div/div/div/iframe"));        
	  driver.switchTo().frame(iframeMsg);
	  WebElement body = driver.findElement(By.cssSelector("body"));
	  Thread.sleep(3000);
	  body.sendKeys("Test submission text final");
	  driver.switchTo().defaultContent();
	  driver.findElement(By.xpath("//input[@name='badgeos_submission_submit']")).click();
	                     
	
}
public void Password(String pr, String dt) throws InterruptedException { 
	driver.findElement(By.xpath(".//*[@id='createuser']/table/tbody/tr[6]/td/button")).click();
	driver.findElement(By.id("pass1-text")).clear();
	//To enter full text in password column
	/* WebElement textbox = driver.findElement(By.name("pass1-text"));
    Actions act = new Actions(driver);
    act.keyUp(Keys.SHIFT).sendKeys(textbox,"Sample@123").keyDown(Keys.SHIFT).build().perform();*/
    driver.findElement(By.id("pass1-text")).clear();

	JavascriptExecutor jse = (JavascriptExecutor) driver;
	Thread.sleep(2000);
	jse.executeScript("document.getElementById('pass1-text').value = 'Sample@12345678910';");  
	
	Thread.sleep(2000);
	/*WebElement chk = driver.findElement(By.cssSelector("input[name='pw_weak']"));
	if(chk.isDisplayed())
	{
		System.out.println("Yes"); 
		chk.click();
	}
	else{
		System.out.println("No");
	}
	*/ 
}

public void logout(String pr, String dt) throws InterruptedException {
	
	
	//Actions action = new Actions(driver);
	//WebElement mainMenu = driver.findElement(By.linkText("Howdy, admin"));
	//System.out.println("dgfdgdfghdf");
	//action.moveToElement(mainMenu).moveToElement(driver.findElement(By.xpath("//a[contains(text(),'Log Out')]"))).click().build().perform();
	//driver.findElement(By.linkText("Howdy, admin")).click();
	//System.out.println("going");
	//driver.findElement(By.xpath("//a[contains(text(),'Log Out')]")).click();
	//driver.findElement(By.linkText("BadgeOS")).click();
	//driver.findElement(By.id("wp-logout")).click();
	//driver.findElement(By.linkText("log out")).click();
	//System.out.println("sdfgfdgdfg");
	//driver.findElement(By.id("wp-logout")).click();
	driver.findElement(By.linkText("BadgeOS")).click();
	 WebElement menu=driver.findElement(By.xpath("/html/body/div/div[2]/div[1]/div/ul[2]/li[1]/a"));
	   WebElement submenu=driver.findElement(By.xpath("//a[contains(text(),'Log Out')]"));
	   Actions action = new Actions(driver);
	   action.moveToElement(menu).perform();
	   submenu.click();
	   
	   //driver.findElement(By.id("wp-logout")).click();
	
}

public void userresult(String pr, String dt) throws InterruptedException, BiffException, IOException {
	
	String[] arrPolpr = pr.split("%"); 
	 
	System.out.println("123");
		 
 	
 	
 	String str=driver.findElement(By.xpath("/html/body/div/div[2]/div[2]/div[1]/div[3]/div/p")).getText();
    //System.out.println(str);
     
    String ele2="New user created. Edit user";
    
    String ele3="ERROR: The email address isn�t correct.";
    
    String ele4="ERROR: This username is already registered. Please choose another one.";
    
    String ele5="ERROR: This email is already registered, please choose another one.";
    
    if(ele2.equals(str))
    { 
     System.out.println("Success message displayed properly as expected");
     writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message","Nill",dt+ "",str+ " - Actual text is displayed", "Pass");
    // Thread.sleep(2000); 
  	
   // objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
    
    //Working     	
  	 File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
     FileUtils.copyFile(scrFile, new File("D:\\Autofiles\\Screenshot\\userCreation.png")); 
 	 driver.close();
  	//Working
  	
    }
    else if(ele3.equals(str)) 
    {
     System.out.println("Email address is not correct");
     writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message","Nill",dt+ "",str+ " - Actual text is displayed", "Fail");
 	// objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
     File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
     FileUtils.copyFile(scrFile, new File("D:\\Autofiles\\Screenshot\\userCreation.png")); 
 	 driver.close();
    }
    else if(ele4.equals(str))
    {
     System.out.println("Username already exists");
     writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message","Nill",dt+ "",str+ " - Actual text is displayed", "Fail");
 	// objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
     File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
     FileUtils.copyFile(scrFile, new File("D:\\Autofiles\\Screenshot\\userCreation.png")); 
 	 driver.close();

    }
    else if(ele5.equals(str))
    {
     System.out.println("Email is already exists");
     writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message","Nill",dt+ "",str+ " - Actual text is displayed", "Fail");
 	// objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
     File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
     FileUtils.copyFile(scrFile, new File("D:\\Autofiles\\Screenshot\\userCreation.png")); 
 	 driver.close();

    }
    else
    {
     System.out.println("Nothing displayed");
     writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message","Nill",dt+ "",str+ " - Actual text is displayed", "Fail");
 	// objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
     File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
     FileUtils.copyFile(scrFile, new File("D:\\Autofiles\\Screenshot\\userCreation.png")); 
 	 driver.close();

    }
    
    driver.close();
      
   }
	
 	/*  	
	if (str.equals(dt)) {
		System.out.println("Ok right");
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	} else {
		System.out.println("Ok wrong"); 
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	}
	*/




public void Del(String pr, String dt) throws InterruptedException {  // Created user deletion
	 driver.findElement(By.id("user-search-input")).sendKeys("AutoUser1234567"); //To change the username
	   driver.findElement(By.id("search-submit")).click();
	   
	   WebElement menu=driver.findElement(By.xpath("//a[contains(text(),'AutoUser1234567')]")); //Change
	   WebElement submenu=driver.findElement(By.className("submitdelete"));
	   Actions action = new Actions(driver);
	   action.moveToElement(menu).perform();
	   submenu.click();
	   
	 //  driver.findElement(By.id("delete_option0")).click();
	   driver.findElement(By.id("submit")).click();  
	   
	
}
	
public void badgeLogin(String pr, String dt) throws InterruptedException {  
	driver.findElement(By.xpath("//input[@name='log']")).sendKeys(new String[] {"admin"});
	driver.findElement(By.xpath("//input[@name='pwd']")).sendKeys(new String[] {"Test@123"});
	driver.findElement(By.id("lwa_wp-submit")).click();    
// driver.findElement(By.id("wp-submit")).click();
   Thread.sleep(3000);    
	
}
  
public void badgeLogout(String pr, String dt) throws InterruptedException {  
	Actions action = new Actions(driver);
	driver.navigate().to("http://dev.asahitechnologies.com/badgeos");
	WebElement mainMenu = driver.findElement(By.xpath("//a[contains(text(),'Howdy, admin')]"));
	action.moveToElement(mainMenu).moveToElement(driver.findElement(By.xpath("//a[contains(text(),'Log Out')]"))).click().build().perform();
	  Thread.sleep(3000);

      driver.navigate().to("http://dev.asahitechnologies.com/badgeos/");   
}
 

public void Parameter(String pr, String dt) throws InterruptedException, BiffException, IOException {  
	
	
	Workbook wb1 = Workbook.getWorkbook(new File(INPUTEXCEL));  
	Sheet sheet1 = wb1.getSheet("Users");
	//System.out.println(sheet1.g);
	int rows1 = sheet1.getRows();
	//int cols1 = sheet1.getColumns();
	
		for (int row1 = 0; row1 <rows1; row1++)  
		{
			//UserName
			System.out.println(sheet1.getCell(1, (row1+1)).getContents());
			String username = sheet1.getCell(1, (row1+1)).getContents();
			//UserName
			 
			//Email
			System.out.println(sheet1.getCell(2, (row1)).getContents());
			String email = sheet1.getCell(2, (row1+1)).getContents();
			//Email
			
			//FirstName
			System.out.println(sheet1.getCell(3, (row1)).getContents());
			String firstname = sheet1.getCell(3, (row1+1)).getContents();
			
			//LastName
			System.out.println(sheet1.getCell(4, (row1)).getContents());
			String Lastname = sheet1.getCell(4, (row1+1)).getContents();
			
		    //URL
			System.out.println(sheet1.getCell(5, (row1)).getContents());
			String Url = sheet1.getCell(5, (row1+1)).getContents();
			
			//Dropdown values
			System.out.println(sheet1.getCell(0, (row1)).getContents());
			String fname = sheet1.getCell(0, (row1+1)).getContents();
			//Dropdown values
						
			
		    driver.findElement(By.id("user_login")).sendKeys(username);
		    driver.findElement(By.id("email")).sendKeys(email);
		    driver.findElement(By.id("first_name")).sendKeys(firstname); 
		    driver.findElement(By.id("last_name")).sendKeys(Lastname);
		    driver.findElement(By.id("url")).sendKeys(Url); 
			driver.findElement(By.id("role")).sendKeys(fname);
			driver.findElement(By.id("createusersub")).click();
			driver.findElement(By.className("page-title-action")).click();
			
		 			
						
 		    }
		
		}
		
			
			
		
			
			/*Select oSelection = new Select(driver.findElement(By.id(pr)));
			List<WebElement> oSize = oSelection.getOptions();
			int iListSize = oSize.size();
			
			String lname = sheet1.getCell((0), (row1)).getContents();
			if (lname==oSize){
				
			}
			System.out.println(lname);			
			//driver.findElement(By.id(pr)).sendKeys(dt);
			//WebElement dd= driver.findElement(By.id(pr));//pr
			//Select mySelect= new Select(dd);
			//mySelect.selectByVisibleText(dt);//dt
		}

	
	*/
	
	 

  


public void achieveResult(String pr, String dt) throws InterruptedException {   
	String[] arrPolpr = pr.split("%");  
	
		
	//String str = driver.findElement(By.xpath("//*[contains(text(),'" +"New user created. Edit user" + "')]")).getText();
	System.out.println("test ok");
	String str =driver.findElement(By.xpath("//strong/a")).getText();  
	//String str =driver.findElement(By.xpath("message/p")).getText();
	System.out.println(str);
	str.toString();
	
	if (str.equals(dt)) {
		System.out.println("Ok right");
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	} else {
		System.out.println("Ok wrong"); 
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	}
	
}
	
public void Dropdown(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(3000); 
	String[] arrPolpr = pr.split("%");
	
	 Select dropdown1 = new Select(driver.findElement(By.id(arrPolpr[0])));
    // dropdown.selectByVisibleText("Submission (Auto-accepted)");
     dropdown1.selectByVisibleText(dt);	
	
	
	
}

public void Dropdownbyxpath(String pr, String dt) throws InterruptedException {
	
	//Thread.sleep(3000); 
	
	 Select dropdown = new Select(driver.findElement(By.xpath(pr))); 
    // dropdown.selectByVisibleText("Submission (Auto-accepted)");
     dropdown.selectByVisibleText(dt);	
	
	
	
}
public void submitText(String pr, String dt) throws InterruptedException {  
	
	  WebElement iframeMsg = driver.findElement(By.xpath("//iframe"));         
	    driver.switchTo().frame(iframeMsg);

	    WebElement body = driver.findElement(By.cssSelector("body"));
	    body.sendKeys("Test submission for achievement"); 
	    driver.switchTo().defaultContent();
	    driver.findElement(By.xpath("//input[@name='badgeos_submission_submit']")).click();
	 
}
public void verifyResultSucessMessage(String pr, String dt) throws InterruptedException {  
	
	 WebElement list1 = driver.findElement(By.xpath("//p"));     
		System.out.println(list1.getText());  
}

public void verifyResultSucessMessage2(String pr, String dt) throws InterruptedException {  
	
	 WebElement list1 = driver.findElement(By.xpath("//p[contains(text(),'Thank you! Your submission has been received.')]"));    
	 list1.toString();
		System.out.println(list1.getText());  
}
public void verifyResultSucessMessage1(String pr, String dt) throws InterruptedException {  
	
	 WebElement list = driver.findElement(By.cssSelector("//span[4]"));   
		System.out.println(list.getText());
}

public void addComments(String pr, String dt) throws InterruptedException {  
	
	
    driver.findElement(By.xpath("//a[contains(text(),'Show/Add Comments')]")).click();
    WebElement iframeMsg1 = driver.findElement(By.xpath("//iframe"));         
    driver.switchTo().frame(iframeMsg1);
    WebElement body1 = driver.findElement(By.cssSelector("body"));
    body1.sendKeys("Submitting comments");
    driver.switchTo().defaultContent(); 
    driver.findElement(By.name("badgeos_comment_submit")).click();
}
 
public void verifyAddcommentsSucessMessage(String pr, String dt) throws InterruptedException {  
	
	
	 WebElement list1 = driver.findElement(By.xpath("//p[contains(text(),'You have earned this achievement!')]"));
	    list1.toString(); 
	    String mystr = "You have earned this achievement!";
	    String str = list1.getText();
	   if(str.equals(mystr))
	   {
		   System.out.println(str);
	   }
	   else 
	   {
		   System.out.println("fail");
		   
	   }
	
}

public void verifyAddcommentsResult(String pr, String dt) throws InterruptedException {
	
	 String[] arrPolpr = pr.split("%"); 
	  
		
		Thread.sleep(2000); 
		 WebElement ele = driver.findElement(By.xpath("//p[contains(text(),'You have earned this achievement!')]"));
		    ele.toString(); 
		 	
	    String str = ele.getText().toString();
	    	
			
		if (str.equals(dt)) {
			System.out.println("Ok AT1");
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("AT! Wrong"); 
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
	
	
	
	
}


public void S1Result(String pr, String dt) throws InterruptedException {  
	 
	 
	 String[] arrPolpr = pr.split("%");  
	  
		
		Thread.sleep(2000);  
		 WebElement ele = driver.findElement(By.xpath("//p[contains(text(),'You have earned this achievement!')]"));
		 // ele.toString(); 
		 	
	    String str = ele.getText().toString();
	    	
			
		if (str.equals(dt)) {
			System.out.println("Pass Score 1");
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("Fail Score 1");  
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
	
	
	}  

public void S2Result(String pr, String dt) throws InterruptedException {  
	  
	  
	 String[] arrPolpr = pr.split("%");  
	   
		 
		Thread.sleep(2000);  
		 WebElement ele = driver.findElement(By.xpath("//p[contains(text(),'Thank you! Your submission has been received.')]"));    
		// ele.toString();
			System.out.println(ele.getText());  
	    	
			 String str = ele.getText().toString();
			 
		if (str.equals(dt)) {
			System.out.println("Pass Score 1");
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("Fail Score 1");  
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
	
	
	}  

 
public void navigateclassname(String pr, String dt) throws InterruptedException { 
	 
	//Thread.sleep(5000); 
	
	WebDriverWait ww = new WebDriverWait(driver, 10);
	ww.until(ExpectedConditions.presenceOfElementLocated(By.className("step-row step-18799")));
		
	 driver.findElement(By.className("step-row step-18799"));
	
	
	//driver.findElement(By.xpath(pr)).click();
	//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	
} 

public void verifyRequiredStep(String pr, String dt) throws InterruptedException { 
	 
	 
	 WebElement list1 = driver.findElement(By.xpath("//h4"));   
		System.out.println(list1.getText());  
	
	
	
	}  
public void verifyS3result(String pr, String dt) throws InterruptedException { 
	 
	 
	 String[] arrPolpr = pr.split("%"); 
	   
		
	 WebElement ele = driver.findElement(By.xpath("//h4"));   
		System.out.println(ele.getText());  
		 	
	    String str = ele.getText().toString(); 
	    	
			 
		if (str.equals(dt)) {
			System.out.println("Pass verifyS3 Result");
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("Fail verifyS3 Result"); 
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
	
	
	}  

	public void clickDeny(String pr, String dt)throws InterruptedException, IOException   {
		
		  
		    
		    JavascriptExecutor jse = (JavascriptExecutor)driver;
		    jse.executeScript("window.scrollBy(0,250)", "");
		    
		    driver.findElement(By.xpath("//a[2]")).click();
		    driver.findElement(By.xpath("//a[contains(text(),'Show/Add Comments')]")).click();
		   
		    JavascriptExecutor jse1 = (JavascriptExecutor)driver;
		    jse1.executeScript("window.scrollBy(0,250)", "");
		    
		    WebElement iframeMsg = driver.findElement(By.xpath("//iframe"));        
		    driver.switchTo().frame(iframeMsg);
		    WebElement body = driver.findElement(By.cssSelector("body")); 
		    body.sendKeys("Your submission Denied by BadgeOS administrator");
		    driver.switchTo().defaultContent();
		    
		    JavascriptExecutor jse3 = (JavascriptExecutor)driver;
		    jse3.executeScript("window.scrollBy(0,250)", "");
		    
		    driver.findElement(By.name("badgeos_comment_submit")).click();		
		
		 
	}
	
	public void DenyResult(String pr, String dt)throws InterruptedException, IOException   {
		 String[] arrPolpr = pr.split("%"); 
		  
			
			Thread.sleep(2000); 
						 	
		    String str = "Denied";

		    	
				
			if (str.equals(dt)) {
				System.out.println("Ok AT1");
			writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
			//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
			} else {
				System.out.println("AT! Wrong"); 
			writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
			//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
			}
				
		
		
		
	}
	
	public void clickApprove(String pr, String dt)throws InterruptedException, IOException   {   
		
		JavascriptExecutor jse4 = (JavascriptExecutor)driver;
	    jse4.executeScript("window.scrollBy(0,250)", "");
	    
	    driver.findElement(By.xpath("//p[2]/a")).click();		
	}
public void ApproveResult(String pr, String dt)throws InterruptedException, IOException   {   
	String[] arrPolpr = pr.split("%"); 
	  
	
	Thread.sleep(2000); 
				 	
    String str = "Approved";

    	
		
	if (str.equals(dt)) {
		System.out.println("Ok AT1");
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	} else {
		System.out.println("AT! Wrong"); 
	writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
	//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
	}
		

			
	} 
	
public void pageVerificationResult(String pr, String dt)throws InterruptedException, IOException   {   
		
	 String[] arrPolpr = pr.split("%"); 
	  
		
		Thread.sleep(2000); 
		String title;
		title = driver.getTitle();
		System.out.println("Page title is: "+title);
	 
				 	
	    String str = title.toString(); 
	    	
			
		if (str.equals(dt)) {
			System.out.println("Ok AT1");
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",str,str+ "",str+ " - Actual text is displayed", "Pass");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		} else {
			System.out.println("AT! Wrong"); 
		writeExcel.writeResult(RESULTEXCEL, arrPolpr[0],arrPolpr[1],"Verify success message",dt,dt+ "",str+ " - Actual text is displayed", "Fail");
		//objDriver.screenshot(driver, arrPolpr[3], arrPolpr[2]);
		}
		
	
		
	}
public void activatePlugins(String pr, String dt)throws InterruptedException, IOException   {   
	
	driver.findElement(By.xpath("//*[@id='wp-admin-bar-site-name']/a")).click(); 
	   
	   driver.findElement(By.xpath("//*[@id='menu-plugins']/a/div[3]")).click();
		
	 String[] chkId1 = { "checkbox_3200d6e9f3e804475f6c613f899baae3","checkbox_5b77597fd43e10542cbf7243308cc742","checkbox_8fd46920b5ea263c50915a994404b7df",
	        "checkbox_33b924bfb3e60b9c281f9f1882a6bcbc",
	        "checkbox_1f55258f0e33c6da3bdc397e8f2c48cb","checkbox_4ed51cf781b052022db2d8f414a53317",
	        "checkbox_2133bda291d6811267e24f014759afc6","checkbox_57dae15ce3c640d7aa209820ddda6fe3",
	        "checkbox_5952a5902dadc674fffd3dbbfdfd5f07"};
	      
	        for(int i=0;i<chkId1.length;i++)
	        {
	         if(i<2)
	         {
	          JavascriptExecutor jse = (JavascriptExecutor)driver;
	          jse.executeScript("window.scrollBy(0,250)", "");
	          
	          driver.findElement(By.id(chkId1[i])).click();
	          WebElement action1=driver.findElement(By.name("action"));
	          
	          JavascriptExecutor jse2 = (JavascriptExecutor)driver;
	          jse2.executeScript("window.scrollBy(0,-250)", "");
	          
	          action1.click();
	          action1.sendKeys(Keys.ARROW_DOWN,Keys.ENTER);
	          
	          WebElement element3 = driver.findElement(By.id("doaction"));
	          
	          JavascriptExecutor executor3 = (JavascriptExecutor)driver;
	          executor3.executeScript("arguments[0].click();", element3);
	          Thread.sleep(3000);
	          
	         }
	         else
	         {
	          JavascriptExecutor jse = (JavascriptExecutor)driver;
	          jse.executeScript("window.scrollBy(0,500)", "");
	          
	          driver.findElement(By.id(chkId1[i])).click();
	          WebElement action1=driver.findElement(By.name("action"));
	          
	          JavascriptExecutor jse2 = (JavascriptExecutor)driver;
	          jse2.executeScript("window.scrollBy(0,-500)", "");
	          
	          action1.click();
	          action1.sendKeys(Keys.ARROW_DOWN,Keys.ENTER);
	          
	          WebElement element3 = driver.findElement(By.id("doaction"));
	          
	          JavascriptExecutor executor3 = (JavascriptExecutor)driver;
	          executor3.executeScript("arguments[0].click();", element3);
	          Thread.sleep(3000);
	         }
	        }
	}
	
public void deActivatePlugins(String pr, String dt)throws InterruptedException, IOException   {   
		 
	driver.findElement(By.xpath("//*[@id='wp-admin-bar-site-name']/a")).click(); 
	   
	   driver.findElement(By.xpath("//*[@id='menu-plugins']/a/div[3]")).click();
	   
	  
	   String[] chkId = {"checkbox_5b77597fd43e10542cbf7243308cc742","checkbox_8fd46920b5ea263c50915a994404b7df",
	     "checkbox_5952a5902dadc674fffd3dbbfdfd5f07","checkbox_33b924bfb3e60b9c281f9f1882a6bcbc",
	     "checkbox_1f55258f0e33c6da3bdc397e8f2c48cb","checkbox_4ed51cf781b052022db2d8f414a53317",
	     "checkbox_2133bda291d6811267e24f014759afc6","checkbox_57dae15ce3c640d7aa209820ddda6fe3",
	     "checkbox_3200d6e9f3e804475f6c613f899baae3"};
	   
	     for(int i=0;i<chkId.length;i++)
	     {
	      if(i<2 || i==8)
	      {
	       JavascriptExecutor jse = (JavascriptExecutor)driver;
	       jse.executeScript("window.scrollBy(0,250)", "");
	       
	       driver.findElement(By.id(chkId[i])).click();
	       WebElement action1=driver.findElement(By.name("action"));
	       
	       JavascriptExecutor jse2 = (JavascriptExecutor)driver;
	       jse2.executeScript("window.scrollBy(0,-250)", "");
	        
	       action1.click();
	       action1.sendKeys(Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.ENTER);
	       
	       WebElement element3 = driver.findElement(By.id("doaction"));
	       
	       JavascriptExecutor executor3 = (JavascriptExecutor)driver;
	       executor3.executeScript("arguments[0].click();", element3);
	       
	      }
	      
	      else 
	      {
	    	  JavascriptExecutor jse = (JavascriptExecutor)driver;
			 jse.executeScript("window.scrollBy(0,500)", "");
			 
			 driver.findElement(By.id(chkId[i])).click();
			 WebElement action1=driver.findElement(By.name("action"));
			 
			 JavascriptExecutor jse2 = (JavascriptExecutor)driver;
			 jse2.executeScript("window.scrollBy(0,-500)", "");
			 
			 action1.click();
			 action1.sendKeys(Keys.ARROW_DOWN,Keys.ARROW_DOWN,Keys.ENTER);
			 
			 WebElement element3 = driver.findElement(By.id("doaction"));
				
				JavascriptExecutor executor3 = (JavascriptExecutor)driver;
				executor3.executeScript("arguments[0].click();", element3);
		 
	 

	
	     }
}


}
}
