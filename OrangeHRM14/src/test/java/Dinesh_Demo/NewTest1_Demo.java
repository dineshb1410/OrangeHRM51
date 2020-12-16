package Dinesh_Demo;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.testng.annotations.Test;

public class NewTest1_Demo {
	
 WebDriver driver;
	
  @Test
  public void f() {
	  
	  
	  
	  
  }
  

   public  WebElement findElement(By by) throws Exception 
   {
			
		   WebElement elem = driver.findElement(by);
		    	    
		    
		    if (driver instanceof JavascriptExecutor) 
		    {
		        ((JavascriptExecutor)driver).executeScript("arguments[0].style.border='3px solid red'", elem);
		        
		  
		        
	 }
		    return elem;
   }
  
}
