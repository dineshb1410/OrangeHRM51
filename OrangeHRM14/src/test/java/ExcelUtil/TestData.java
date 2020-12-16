package ExcelUtil;

import org.testng.annotations.Test;

public class TestData {
	
	
  @Test
  public void MyTestData() {
	  
	  int Rows=3,Columns=3;
	  
	  String[][] mydata=new String[Rows][Columns];
  	
  		mydata[0][0]="UseName";
    	mydata[0][1]="Password";
    	mydata[0][2]="Nationality Text";
  	
    	mydata[1][0]="Admin1";
    	mydata[1][1]="admin1";
    	mydata[1][2]="Indian1";
    	
    	mydata[2][0]="Admin2";
    	mydata[2][1]="admin2";
    	mydata[2][2]="Indian2";
    	
    	for(int i=0;i<=Rows-1;i++)
    	{
    		for(int j=0;j<=Columns-1;j++)
    		{
    			System.out.print(mydata[i][j]+"\t");
    		}
    		System.out.println("");
    	}
  	
  }
  
  
  
}
