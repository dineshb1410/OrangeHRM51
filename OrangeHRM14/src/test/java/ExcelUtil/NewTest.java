package ExcelUtil;

import org.testng.annotations.Test;

import ExcelUtil.ExcelApiTest6;

public class NewTest {
	
	
	
  @Test
  public void getMyxcelata() throws  Exception {
	  
	  ExcelApiTest6 eat = new ExcelApiTest6();
	  String[][] mydata = eat.getTableArray1("C://HTML Report//OrangeHRM6//TC01_Nationality1.xlsx", "Sheet1");
	  
		System.out.println(mydata[3][3]);
  }
  
  
  
  
}
