package driverScript;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunLibrary.FunctionLibrary;
import utilities.ExcelFileUtil;
import utilities.PropertyFileUtil;

public class DriverScript {
	static ExtentReports report;
	static ExtentTest test;
	public static WebDriver driver;

	public static void main(String[] args) throws Exception {
		ExcelFileUtil excel=new ExcelFileUtil();
		for (int i = 1; i < excel.rowCount("MasterTestCases") ; i++) {
			String ModuleStatus="";
			if (excel.getData("MasterTestCases", i, 2).equalsIgnoreCase("Y")){
				String TCModule=excel.getData("MasterTestCases", i, 1);
				report=new ExtentReports("D:\\RAJIB\\Selenium\\Hybrid\\Reports\\"+TCModule+FunctionLibrary.generateDate());
				test=report.startTest(TCModule);
				for (int j = 1; j <excel.rowCount(TCModule); j++){
					String Description=excel.getData(TCModule, j, 0);
					String Function_Name=excel.getData(TCModule, j, 1);
					String Locator_Type=excel.getData(TCModule, j, 2);
					String Locator_Value=excel.getData(TCModule, j, 3);
					String Test_Data=excel.getData(TCModule, j, 4);
					
					System.out.println(Description+" "+Function_Name+" "+Locator_Type+" "+Locator_Value+" "+Test_Data);
					
					try{
						if(Function_Name.equalsIgnoreCase("startBrowser")){
							
						      driver=	FunctionLibrary.startBrowser();
						      test.log(LogStatus.INFO, Description);
						}else if(Function_Name.equalsIgnoreCase("openApplication")){
							 FunctionLibrary.openApplication(driver);
							 test.log(LogStatus.INFO, Description);
						}else if(Function_Name.equalsIgnoreCase("waitForElement")){
							FunctionLibrary.waitForElement(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
						}else if(Function_Name.equalsIgnoreCase("typeAction")){
							FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
							test.log(LogStatus.INFO, Description);
						}else if(Function_Name.equalsIgnoreCase("clickAction")){
							FunctionLibrary.clickAction(driver, Locator_Type, Locator_Value);
							test.log(LogStatus.INFO, Description);
						}else if(Function_Name.equalsIgnoreCase("closeBrowser")){
							FunctionLibrary.closeBrowser(driver);
							test.log(LogStatus.INFO, Description);
						}
						excel.setData(TCModule, j, 5, "PASS");
						ModuleStatus="PASS";
						test.log(LogStatus.PASS, Description);
						
					}catch(Exception e){
						excel.setData(TCModule, j, 5, "FAIL");
						ModuleStatus="FAIL";
						String reqDate=FunctionLibrary.generateDate();
						File srcFile=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						FileUtils.copyFile(srcFile, new File("D:\\RAJIB\\Selenium\\Hybrid\\Screenshots"+Description+FunctionLibrary.generateDate()+".png"));
						test.log(LogStatus.FAIL, Description);
						test.log(LogStatus.INFO,test.addScreenCapture("D:\\RAJIB\\Selenium\\Hybrid\\Screenshots"+Description+FunctionLibrary.generateDate()+".png"));
						break;
					}		
				}
				
				if(ModuleStatus.equalsIgnoreCase("PASS")){
					excel.setData("MasterTestCases", i, 3, "PASS");
				}else{
					excel.setData("MasterTestCases", i, 3, "FAIL");
				}
				
				report.endTest(test);
               report.flush();
				
			}else{
				excel.setData("MasterTestCases", i, 3, "Not Executed");
			}
			
			
		}

	}

}

					
					
					

