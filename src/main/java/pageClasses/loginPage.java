package pageClasses;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import baseClass.pageBaseClass;

public class loginPage extends pageBaseClass{
	
	public loginPage(WebDriver driver) {
		this.driver = driver;
	}
	
	public homePage loginAccount() throws IOException {
		
		//Reading Credentials from excel file
		File src = new File(System.getProperty("user.dir") + "\\Excel\\data.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//Used to format data from number to string
		DataFormatter formatter = new DataFormatter();
		
		String email = formatter.formatCellValue(sheet.getRow(1).getCell(0));
		String password = formatter.formatCellValue(sheet.getRow(2).getCell(0));
		
		//Closed the workbook
		workbook.close();
		
		
		//Explicit wait
		WebDriverWait wait = new WebDriverWait(driver, 20);
		try {
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(prop.getProperty("login_xpath")))));
		} catch (StaleElementReferenceException ex) {
			// TODO: handle exception
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(prop.getProperty("login_xpath")))));
		}
		
		//Entered ID
		driver.findElement(By.xpath(prop.getProperty("login_xpath"))).sendKeys(email);
		reportPass("Email Id entered Successfully");
		screenshot("email.png", driver);
		driver.findElement(By.id(prop.getProperty("nextBtn_id"))).click();
		
		try {
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(prop.getProperty("password_xpath")))));
		} catch (StaleElementReferenceException ex) {
			// TODO: handle exception
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(prop.getProperty("password_xpath")))));
		}
		
		//Entered Password
		driver.findElement(By.xpath(prop.getProperty("password_xpath"))).sendKeys(password);
		reportPass("Password entered Successfully");
		screenshot("password.png", driver);
		driver.findElement(By.id(prop.getProperty("nextBtn_id"))).click();
		

		try {
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(prop.getProperty("verifyText_xpath")))));
		} catch (StaleElementReferenceException ex) {
			// TODO: handle exception
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath(prop.getProperty("verifyText_xpath")))));
		}
		

		try {
			driver.findElement(By.id(prop.getProperty("nextBtn_id"))).click();
		} catch (StaleElementReferenceException ex) {
			// TODO: handle exception
			driver.findElement(By.id(prop.getProperty("nextBtn_id"))).click();
		}
		
		
		return PageFactory.initElements(driver, homePage.class);
	}
}
