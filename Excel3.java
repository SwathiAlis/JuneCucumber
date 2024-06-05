package first;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Excel3 {
	public static WebDriver driver;


	public static void main(String[] args) throws Throwable  {
		WebDriverManager.edgedriver().setup();
		EdgeOptions options = new EdgeOptions();
		driver = new EdgeDriver(options);
		driver.get("https://www.makemytrip.com/");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

		driver.findElement(By.xpath("//span[@class='commonModal__close']")).click();
		
		driver.findElement(By.xpath("//span[text()='Buses' and @class='headerIconTextAlignment chNavText darkGreyText']")).click();
		
		driver.findElement(By.xpath("//input[@id='fromCity']")).click();
		driver.findElement(By.xpath("//input[@placeholder='From']")).sendKeys("Chennai");
		driver.findElement(By.xpath("//span[text()='Chennai, Tamil Nadu']")).click();
		
		driver.findElement(By.xpath("//input[@placeholder='To']")).sendKeys("Madurai");
		driver.findElement(By.xpath("//span[text()='Madurai, Tamil Nadu']")).click();
		
		driver.findElement(By.xpath("//div[@aria-label='Thu Jun 06 2024']")).click();
		driver.findElement(By.xpath("//button[@id='search_button']")).click();

		File f = new File ("C:\\Users\\Keerthana\\Desktop\\Buses.xlsx");
		FileOutputStream f1 = new FileOutputStream(f);
		XSSFWorkbook work = new XSSFWorkbook();
		XSSFSheet sheet = work.createSheet("Bus Lists");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		

		List<WebElement> element= driver.findElements(By.xpath("//p[contains(@class,'makeFlex hrtlCenter appendBottom')]"));
		List<WebElement> element1 = driver.findElements(By.xpath("//span[contains(@class,'latoBlack blackText')]"));
		List<WebElement> element2 = driver.findElements(By.xpath("//span[contains(@class,'latoRegular')]"));
		List<WebElement> element3 = driver.findElements(By.xpath("//span[@id='price']"));
		for (int i = 0; i < element.size(); i++) {
			String text = element.get(i).getText();
			String text1 = element1.get(i).getText();
			String text2 = element2.get(i).getText();
			String text3 = element3.get(i).getText();
			System.out.println("Bus Name : " +text+ "Departure Time : " +text1+ "Arrival Time : " +text2+ "Bus Fare : " +text3);
		
			XSSFRow row1 = sheet.createRow(i);
			XSSFCell cell1 = row1.createCell(0);
			cell1.setCellValue(text);
			
			XSSFCell cell2 = row1.createCell(1);
			cell2.setCellValue(text1);
		}	
				
				work.write(f1);
				f1.close();
				
				driver.findElement(By.xpath("(//div[@class='sc-jKJlTe fnCpOO' or @data-test-id='select-seats'])[last()]")).click();
				driver.findElement(By.xpath("(//span[@data-testid='seat_horizontal_sleeper_available' or @class='listingSprite commonSmallSeatIcon appendBottom4'])[10]")).click();
				driver.findElement(By.xpath("//span[text()='Continue']")).click();
				
				driver.findElement(By.xpath("//input[@class='font14 width300']")).sendKeys("Swathi");
				driver.findElement(By.xpath("//input[@class='font14 width101']")).sendKeys("27");
				driver.findElement(By.xpath("//div[text()='Female']")).click();
				driver.findElement(By.xpath("//input[@id='dt_state_gst_info']")).click();
				driver.findElement(By.xpath("//li[text()='Himachal Pradesh']")).click();
				driver.findElement(By.xpath("//p[text()='Confirm and save billing details to your profile']")).click();
				driver.findElement(By.xpath("//input[@type='email']")).sendKeys("swathialis@gmail.com");
				driver.findElement(By.xpath("//input[@name='Mobile Number']")).sendKeys("9094075764");
				
				JavascriptExecutor js = (JavascriptExecutor)driver;
				WebElement up = driver.findElement(By.xpath("//span[text()='Continue']"));
				js.executeScript("arguments[0].scrollIntoView(false)",up);
				Thread.sleep(3000);
				driver.findElement(By.xpath("//span[text()='Continue']")).click();
				driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
				
//			
	
	}
	
}
