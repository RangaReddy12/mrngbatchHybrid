package commonFunctions;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import utilities.PropertyFileUtil;

public class FunctionLibrary {
	public static WebDriver driver;
	public static String Expected="";
	public static String Actual ="";
	//method or launch browser
	public static WebDriver startBrowser()throws Throwable
	{
		if(PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("chrome"))
		{
			driver = new ChromeDriver();
			driver.manage().window().maximize();

		}
		else if(PropertyFileUtil.getValueForKey("Browser").equalsIgnoreCase("firefox"))
		{
			driver = new FirefoxDriver();
		}
		else
		{
			System.out.println("Browser value is not matching");
		}
		return driver;
	}
	//method for launch url
	public static void openApplication(WebDriver driver)throws Throwable
	{
		driver.get(PropertyFileUtil.getValueForKey("Url"));
	}
	//method for wait for element
	public static void waitForElement(WebDriver driver,String Locator_Type,String LocatorValue,String Testdata)
	{
		WebDriverWait  myWait = new WebDriverWait(driver, Integer.parseInt(Testdata));
		if(Locator_Type.equalsIgnoreCase("name"))
		{
			myWait.until(ExpectedConditions.visibilityOfElementLocated(By.name(LocatorValue)));
		}
		else if(Locator_Type.equalsIgnoreCase("xpath"))
		{
			myWait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(LocatorValue)));
		}
		else if(Locator_Type.equalsIgnoreCase("id"))
		{
			myWait.until(ExpectedConditions.visibilityOfElementLocated(By.id(LocatorValue)));
		}
	}
	//method for type action
	public static void typeAction(WebDriver driver,String Locator_Type,String Locator_Value,String TestData)
	{
		if(Locator_Type.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(Locator_Value)).clear();
			driver.findElement(By.name(Locator_Value)).sendKeys(TestData);
		}
		else if(Locator_Type.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(Locator_Value)).clear();
			driver.findElement(By.xpath(Locator_Value)).sendKeys(TestData);
		}
		else if(Locator_Type.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(Locator_Value)).clear();
			driver.findElement(By.id(Locator_Value)).sendKeys(TestData);
		}
	}
	//method for button click
	public static void clickAction(WebDriver driver,String Locator_Type,String Locator_Value)
	{
		if(Locator_Type.equalsIgnoreCase("name"))
		{
			driver.findElement(By.name(Locator_Value)).click();
		}
		else if(Locator_Type.equalsIgnoreCase("id"))
		{
			driver.findElement(By.id(Locator_Value)).sendKeys(Keys.ENTER);
		}
		else if(Locator_Type.equalsIgnoreCase("xpath"))
		{
			driver.findElement(By.xpath(Locator_Value)).click();
		}
	}
	//method for validate title
	public static void ValidateTitle(WebDriver driver,String Expected_Title)
	{
		String Actual_Title =driver.getTitle();
		try {
			Assert.assertEquals(Actual_Title, Expected_Title,"Title is Not Matching");
		}catch(Throwable t)
		{
			System.out.println(t.getMessage());
		}
	}
	//method for close bowser
	public static void closeBrowser(WebDriver driver)
	{
		driver.close();
	}
	//method for mouse click
	public static void mouseClick(WebDriver driver) throws Throwable
	{
		Actions ac = new Actions(driver);
		ac.moveToElement(driver.findElement(By.xpath("//a[.='Stock Items ']"))).perform();
		Thread.sleep(3000);
		ac.moveToElement(driver.findElement(By.xpath("(//a)[84]"))).click().perform();
	}
	//method for stock table validation
	public static void categoryTable(WebDriver driver,String ExpectedData)throws Throwable
	{
		if(!driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-textbox"))).isDisplayed())
			Thread.sleep(3000);
		//click search panel 
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-panel"))).click();
		Thread.sleep(3000);
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-textbox"))).sendKeys(ExpectedData);
		Thread.sleep(3000);
		driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-button"))).click();
		Thread.sleep(3000);
		//capture category name from table
		String ActualData =driver.findElement(By.xpath("//table[@id='tbl_a_stock_categorieslist']/tbody/tr[1]/td[4]/div/span/span")).getText();
		System.out.println(ExpectedData+"      "+ActualData);
		Assert.assertEquals(ExpectedData, ActualData,"Categaroy Name Not Found in Table");
	}
	
	//method for capture data
		public static void captureData(WebDriver driver, String LocatorType,String LocatorValue)
		{
			Expected =driver.findElement(By.name(LocatorValue)).getAttribute("value");
		}
		//method for supplier table
		public static void supplierTable(WebDriver driver)throws Throwable
		{
			if(!driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-textbox"))).isDisplayed())
				Thread.sleep(3000);
			//click search panel 
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-panel"))).click();
			Thread.sleep(3000);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-textbox"))).sendKeys(Expected);
			Thread.sleep(3000);
			driver.findElement(By.xpath(PropertyFileUtil.getValueForKey("search-button"))).click();
			Thread.sleep(3000);
		 Actual =driver.findElement(By.xpath("//table[@id='tbl_a_supplierslist']/tbody/tr[1]/td[6]/div/span/span")).getText();
			System.out.println(Expected+"       "+Actual);
			Assert.assertEquals(Expected, Actual,"Supplier Number Not Found in table");
			
		}
		//method for java time stamp
		public static String generateDate()
		{
			Date date = new Date();
			DateFormat df = new SimpleDateFormat("YYYY_MM_dd");
			return df.format(date);
		}
		public static void add()
		{
			int a=876,b=9,c;
			c=a+b;
			System.out.println(c);
		}
		public static void div()
		{
			int a=876,b=9,c;
			c=a+b;
			System.out.println(c);
		}
		public static void mul()
		{
			int a=534543534,b=976765765,c;
			c=a+b;
			System.out.println(c);
		}
	}










