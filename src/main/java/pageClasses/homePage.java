package pageClasses;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import baseClass.pageBaseClass;

public class homePage extends pageBaseClass {

	public WebDriver driver;

	// Constructor
	public homePage(WebDriver driver) {
		this.driver = driver;
	}

	public void writeExcel() throws IOException {

		
		
		File src = new File(System.getProperty("user.dir") + "\\Excel\\data.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		
		String imgSRC;
		driver.manage().timeouts().implicitlyWait(100, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(100, TimeUnit.SECONDS);

		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,250)");

		//Written "Resume Learning" to Excel Sheet
		driver.findElement(By.xpath(prop.getProperty("slider_xpath"))).click();
		List<WebElement> resumeLearningP2 = driver.findElements(By.xpath(prop.getProperty("resumeLearningList_xpath")));
		int n2 = resumeLearningP2.size();
		for (int i = 0; i < n2; i++) {
			//System.out.println(resumeLearningP2.get(i).getText());
			sheet.getRow(i+1).createCell(1).setCellValue(resumeLearningP2.get(i).getText());
		}try {Thread.sleep(2000);}catch (InterruptedException e) {e.printStackTrace();}
		
		//Switching to iframe as banner is on iframe
		driver.switchTo().frame(1);

		//Writing banner data to Excel
		for (int i = 1; i <= 3; i++) {
			imgSRC = driver.findElement(By.xpath(prop.getProperty("beforeSlideCard_xpath1") + i + prop.getProperty("beforeSlideCard_xpath2"))).getAttribute("src");
			if (imgSRC.equalsIgnoreCase(prop.getProperty("img" + i))) {
				System.out.println(prop.getProperty("imtxt" + i));
				sheet.getRow(i).createCell(2).setCellValue(prop.getProperty("imtxt" + i));
			}
		}
		
		
		//Writing remaining banner data to Excel
		int j=0;
		for (int i = 1; i <= 3; i++) {
			j=i+3;
			imgSRC = driver.findElement(By.xpath(prop.getProperty("afterSlideCard_xpath1") + i + prop.getProperty("afterSlideCard_xpath2"))).getAttribute("src");
			if (imgSRC.equalsIgnoreCase(prop.getProperty("img" + j))) {
				System.out.println(prop.getProperty("imtxt" + j));
				sheet.getRow(j).createCell(2).setCellValue(prop.getProperty("imtxt" + j));
			}
		}
		FileOutputStream fout = new FileOutputStream(src);
		workbook.write(fout);
		workbook.close();
		reportPass("Written Banner text to Excel Sheet Successfully");
		reportPass("Written Resume Learning to Excel Sheet Successfully");
		screenshot("homepage.png", driver);
		//Closing driver
		driver.quit();
	}
}
