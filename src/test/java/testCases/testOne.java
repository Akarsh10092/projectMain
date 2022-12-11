package testCases;

import java.io.IOException;

import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import baseClass.pageBaseClass;
import pageClasses.homePage;
import pageClasses.loginPage;

public class testOne extends pageBaseClass{
	homePage homePage;
	loginPage loginPage;
	
	@Parameters("browser")
	@Test
	public void writeExcelTest(String browser) throws IOException {
		
		
		
		//Extent report initialization
		logger = report.createTest("Clearn Floating Data and Resume Learning Text");
		pageBaseClass pageBaseClass = new pageBaseClass();
		pageBaseClass.invokeBrowser(browser);
		loginPage = pageBaseClass.OpenApplication();
		homePage = loginPage.loginAccount();
		homePage.writeExcel();
	}
}
