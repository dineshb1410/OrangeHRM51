package dec16;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class NewTest_login_16 {
	WebDriver driver;

  @Test
  public void Skill_Test() throws Exception {
	  
	  NewTest_login_16 T1 = new NewTest_login_16();
			  T1.OpenChromeBrowser();
			  T1.OpenOrangeHRM();
			  T1.Login();
			  T1.AddSkills();
	  
  }
  
  public void OpenChromeBrowser()  throws Exception {
	  
		System.setProperty("webdriver.chrome.driver","C:\\chromedriver_win32\\chromedriver.exe");
		driver =new ChromeDriver();
		driver.manage().window().maximize() ;	
  }
		
	public  void  OpenOrangeHRM() throws Exception{
		
		driver.get("https://opensource-demo.orangehrmlive.com/");
		
		}
	public  void Login() throws Exception{
		
		driver.findElement(By.xpath("//*[@id='txtUsername']")).sendKeys("Admin");
		driver.findElement(By.name("txtPassword")).sendKeys("admin123");
		driver.findElement(By.name("Submit")).click();
		
	}
	public  void AddSkills() throws Exception{
		
		driver.findElement(By.linkText("Admin")).click();
		driver.findElement(By.partialLinkText("Qualifica")).click();
		driver.findElement(By.id("menu_admin_viewSkills")).click();
		driver.findElement(By.id("btnAdd")).click();
		driver.findElement(By.name("skill[name]")).sendKeys("java1111");
		driver.findElement(By.name("skill[description]")).sendKeys("java123");
		driver.findElement(By.id("btnSave")).click();
	}
	
}
