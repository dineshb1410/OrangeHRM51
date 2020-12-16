package TC02_NaukriTest;

import org.testng.annotations.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import Day_005_TestNG_DataProvider_Lab1.TC03_Login_Static_Paarameters1;

public class NewTest {
  
  static WebDriver driver;
	
	

	 @DataProvider(name = "TC01_OrangeHRM")
	  public static Object[][] TestDataforTest() throws Exception {
		 
	        return new Object[][] { 
	        	
	        	{ "https://www.naukri.com/","Ramesh","Ramesh@gmail.com","Ramesh123","1234567890","4","6" }
	        	
	        	
	        	};
	  }
	 
	 

	@Test(dataProvider="TC01_OrangeHRM")
	public void Login_Test(String TestURL,String Name1,String Email1,String Password1,String MobileNo1,String TypeYearsOfExp1,String TypeMonths1) throws Exception
	{
		
		NewTest.OpenChromeBrowser();
		NewTest.OpenNaukri(TestURL);
		NewTest.LoginNaukri();
		NewTest.AddDetails(Name1 ,Email1 ,Password1 ,MobileNo1 , TypeYearsOfExp1 ,TypeMonths1);
		driver.quit();
	}
  
  public static void LoginNaukri() throws Exception
	{   
	    findElement(By.xpath("//*[@id=\'login_Layer\']/div")).click();
		findElement(By.xpath("//*[@id=\'root\']/div[2]/div[2]/div/form/div[8]/a")).click();
		findElement(By.name("userType")).click();
		
	}
	
	public static void OpenNaukri(String TestURL1) throws Exception
	{
		
		driver.get(TestURL1);
	}
	
	
	public static  WebElement findElement(By by) throws Exception 
	{
				
		 WebElement elem = driver.findElement(by);    	    
		
		 
		if (driver instanceof JavascriptExecutor) 
		{
		 ((JavascriptExecutor)driver).executeScript("arguments[0].style.border='3px solid red'", elem);
	 
		}
		
		return elem;
	}

	
	
	
	
	
	
	
	
	public static void OpenChromeBrowser() throws Exception
	{
		System.setProperty("webdriver.chrome.driver","C:\\chromedriver_win32\\chromedriver.exe");
		driver =new ChromeDriver();
		driver.manage().window().maximize() ;	
	
	}
	
	
	
	
	
	public static   void AddDetails(String Name1,String Email1,String Password1,String MobileNo1,String TypeYearsOfExp1,String TypeMonths1) throws Exception
	{
		

		findElement(By.id("fname")).sendKeys(Name1);
		findElement(By.xpath("//*[@id=\'email\']")).sendKeys(Email1);
		findElement(By.name("password")).sendKeys(Password1);
		findElement(By.name("number")).sendKeys(MobileNo1);
       
		Select ExpYears = new Select(driver.findElement(By.name("expYear")));
		ExpYears.selectByVisibleText(TypeYearsOfExp1);
		
		Select ExpMonth = new Select(driver.findElement(By.name("expMonth")));
		ExpYears.selectByVisibleText(TypeMonths1);
		
		
		
		
	}
	
}
