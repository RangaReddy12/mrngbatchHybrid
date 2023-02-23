package driverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtil;

public class DriverScript {
	public static WebDriver driver;
String inputpath ="D:\\MrngBatch_OJT\\StockAccounting_ERP\\FileInput\\DataEngine.xlsx";
String outputpath ="D:\\MrngBatch_OJT\\StockAccounting_ERP\\FileOutput\\HybridResults.xlsx";
ExtentReports reports;
ExtentTest test;
public void startTest()throws Throwable
{
	String ModuleStatus ="";
	///create reference object
	ExcelFileUtil xl = new ExcelFileUtil(inputpath);
	//iterate all rows in master test cases sheet
	for(int i=1;i<=xl.rowCount("MasterTestCases");i++)
	{
		if(xl.getCellData("MasterTestCases", i, 2).equalsIgnoreCase("Y"))
		{
			//store correspoding sheet into TCModule variable
			String TCModule =xl.getCellData("MasterTestCases", i, 1);
			reports = new ExtentReports("./ExtentReports/"+TCModule+FunctionLibrary.generateDate()+".html");
			//iterate all rows in TCModule sheet
			for(int j=1;j<=xl.rowCount(TCModule);j++)
			{
				
				String Description =xl.getCellData(TCModule, j, 0);
				String Object_Type = xl.getCellData(TCModule, j, 1);
				String Locator_Type = xl.getCellData(TCModule, j, 2);
				String Locator_Value = xl.getCellData(TCModule, j, 3);
				String TestData = xl.getCellData(TCModule, j, 4);
				
				test= reports.startTest(TCModule);
				test.assignAuthor("Ranga");
				test.assignCategory("Functional");
				try {
					if(Object_Type.equalsIgnoreCase("startBrowser"))
					{
						driver =FunctionLibrary.startBrowser();
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("openApplication"))
					{
						FunctionLibrary.openApplication(driver);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("waitForElement"))
					{
						FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, TestData);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("typeAction"))
					{
						FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, TestData);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("clickAction"))
					{
						FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("ValidateTitle"))
					{
						FunctionLibrary.ValidateTitle(driver, TestData);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("closeBrowser"))
					{
						FunctionLibrary.closeBrowser(driver);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("mouseClick"))
					{
						FunctionLibrary.mouseClick(driver);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("categoryTable"))
					{
						FunctionLibrary.categoryTable(driver, TestData);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("captureData"))
					{
						FunctionLibrary.captureData(driver, Locator_Type, Locator_Value);
						test.log(LogStatus.INFO, Description);
					}
					else if(Object_Type.equalsIgnoreCase("supplierTable"))
					{
						FunctionLibrary.supplierTable(driver);
						test.log(LogStatus.INFO, Description);
					}
					//write as pass in status cell
					xl.setCellData(TCModule, j, 5, "Pass", outputpath);
					test.log(LogStatus.PASS, Description);
					ModuleStatus="True";
				}catch(Exception e)
				{
					System.out.println(e.getMessage());
					//write as Fail in status cell
					xl.setCellData(TCModule, j, 5, "Fail", outputpath);
					test.log(LogStatus.FAIL, Description);
					ModuleStatus="False";
					File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					
					FileUtils.copyFile(scrFile, new File("./Screenshots/"+Description+"_"+FunctionLibrary.generateDate()+".png"));
					
					String image = test.addScreenCapture("./Screenshots/"+Description+"_"+FunctionLibrary.generateDate()+".png");
					
					test.log(LogStatus.FAIL, image);
					break;		
				}
				catch(AssertionError e)
				{
					xl.setCellData(TCModule, j, 5, "Fail", outputpath);
					ModuleStatus = "false";
					test.log(LogStatus.FAIL, Description + "Fail");
					
					File scrFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
                    
					FileUtils.copyFile(scrFile, new File("./Screenshots/"+Description+"_"+FunctionLibrary.generateDate()+".jpg"));
					
					String image = test.addScreenCapture("./Screenshots/"+Description+"_"+FunctionLibrary.generateDate()+".jpg");
					
					test.log(LogStatus.FAIL, image);
					
					break;
				}
				 if(ModuleStatus.equalsIgnoreCase("True"))
				 {
					xl.setCellData("MasterTestCases", i, 3, "Pass", outputpath); 
				 }
				 else
				 {
					 xl.setCellData("MasterTestCases", i, 3, "Fail", outputpath);
				 }
			}
			reports.endTest(test);
			reports.flush();
			
		}
		else
		{
			//which test case flag to N Write as Blocked into status cell
			xl.setCellData("MasterTestCases", i, 3, "Blocked", outputpath);
		}
	}
	
}
}
