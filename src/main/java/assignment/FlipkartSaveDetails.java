package assignment;

import java.awt.AWTException;



import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import java.util.ArrayList;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FlipkartSaveDetails {
private WebDriver driver;
Actions builder;
	
	@BeforeClass
	void setup() throws AWTException
	{
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver(); 
		builder = new Actions(driver); 
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.MILLISECONDS);
		
	}
	@AfterClass
	void close()
	{
		driver.close(); 
	}
	
	/*
	 * Retrieving details of the Iphones on Flipkart 
	 */
	
	@Test
	void reteriveIphoneDetails() throws IOException 
	{
		
		String url = "https://www.flipkart.com"; 
		driver.get(url); //Calling the url
		driver.findElement(By.xpath("/html[1]/body[1]/div[2]/div[1]/div[1]/button[1]")).click();//Identifying close icon and clicking it
		
		WebElement searchBox = driver.findElement(By.name("q"));//Identifying search box 
		searchBox.sendKeys("Iphone");// Sending data to search
		searchBox.sendKeys(Keys.ENTER);// Pressing Enter Key
		
		driver.manage().timeouts().implicitlyWait(20,TimeUnit.MILLISECONDS);
		WebDriverWait wait = new WebDriverWait(driver,10);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.
				xpath("/html[1]/body[1]/div[1]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/section[2]/div[4]/div[3]/select[1]")));
		Select Maxprice = new Select(driver.findElement(By.
				xpath("/html[1]/body[1]/div[1]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/section[2]/div[4]/div[3]/select[1]")));
		Maxprice.selectByIndex(9);// Selecting the maximum price
		List <WebElement> product = driver.findElements(By.
				xpath("//div[@class='_3wU53n']"));// Identifying and storing Product name
		List <WebElement> price = driver.findElements(By.
				xpath("//div[@class='_1vC4OE _2rQ-NK']"));// Identifying and storing Product price
		List <WebElement> reviews = driver.findElements(By.className("_38sUEc"));// Identifying and storing Product reviews

		
		List<ArrayList<String>> listOfIphonesDetails = new ArrayList<ArrayList<String>>();
		for(int i=0;i<product.size();i++) {
			try { //Iterating and saving the details of each product
				ArrayList<String> items = new ArrayList<String>();

				items.add(product.get(i).getText());
				items.add(reviews.get(i).getText());
			
				String productPriceText = price.get(i).getText().replaceAll("[^0-9]", "");
				int productPrice = Integer.parseInt(productPriceText); //Converting type of price into int
			
				if(productPrice > 40000) { // Condition to check price of product less than 40000 
					continue;
				}
				items.add(productPriceText);
				listOfIphonesDetails.add(items);
				writeToExcelFile(listOfIphonesDetails);
			}catch(Exception e) {
				System.out.println("Exception : " + e);
			}	
		}
		
	}
	/*
	 * Writing details of Iphones in the Excel file
	 */
	private void writeToExcelFile(List<ArrayList<String>> listOfIphonesDetails) {
		
		try{
			FileOutputStream file = new FileOutputStream("IphoneDetails.xlsx");
		
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("Less than 40K");
		
			for(int i = 0;i<listOfIphonesDetails.size();i++) { // Writing the retrieved data in excel file
			
				ArrayList<String> temp = listOfIphonesDetails.get(i);
				HSSFRow row1 = sheet.createRow(i);
				row1.createCell(0).setCellValue(new HSSFRichTextString(temp.get(0)));
				row1.createCell(1).setCellValue(new HSSFRichTextString(temp.get(1)));
				row1.createCell(2).setCellValue(new HSSFRichTextString(temp.get(2)));
			}
		
			workbook.write(file);
			workbook.close();
			file.close(); //close the excel file
		}
		catch(Exception e) {
			System.out.println("File operation failed : " + e);
		}
		
	}
}
