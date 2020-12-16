package ExcelUtil;

import org.testng.annotations.Test;

public class TC01 {
	
	@Test
	public void Test1()throws Exception
	{
		

    	ExcelApiTest5 eat=new ExcelApiTest5();
    	
    	eat.setAutoFilter_Single1("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet2",7,7L,"Apple");
    	
    	Thread.sleep(5000);
    	
		String str1=eat.getCellData("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet3",0,0);
    	System.out.println("Sum of Values "+str1);
    	
	}

}
