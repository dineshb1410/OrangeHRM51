package ExcelUtil;

import org.apache.poi.ss.util.CellRangeAddress;

/*import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
*/




import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.common.usermodel.HyperlinkType;
//import org.apache.poi.xssf.usermodel.XSSFDateUtil;
import org.apache.poi.ss.usermodel.*;


public  class ExcelApiTest5
{
   public FileInputStream fis = null;
    public XSSFWorkbook workbook = null;
    public XSSFSheet sheet = null;
    public XSSFRow row = null;
    public XSSFCell cell = null;
    public FileOutputStream fout=null;
    

    
    public static void main (String args[]) throws Exception
    {
    	
    		
    	ExcelApiTest5 eat=new ExcelApiTest5();
    	
    	//eat.setAutoFilter_Multiple("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet8",2,"admin124");
      
    //	eat.setAutoFilter_Single1("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet8",2,2L,"Linda");
        
    //eat.setAutoFilter_Single1("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet2",4,4L,"Lee");

    eat.setAutoFilter_Single1("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet2",7,7L,"Apple");
    	
    	//eat.setAutoFilter_Single1("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet2",7,7L,"Apple");
  //  eat.setAutoFilter_Single1_Sum("C://HTML Report//OrangeHRM6//TC01_AddEmp2.xlsx","Sheet2",7,7L,"Apple");
    	
    	//Thread.sleep(5000);
    	

    //	String str=eat.getCellData("C://HTML Report//OrangeHRM6//TC01_AddEmp3.xlsx","Sheet3",0,1);
    //	System.out.println("Sum of Values "+str);
    
    	
    	
    	//String str1=eat.getCellData("C://HTML Report//OrangeHRM6//TC01_AddEmp3.xlsx","Sheet3",0,0);
    //	System.out.println("Sum of Values "+str1);
    	
    	
    	//eat=null;
    	//System.gc();
    	
    	
    	 	
    	
    }
    
    
    
    
public  void setAutoFilter_Single1(String xlFilePath,String sheetName, int columnIndex,long colId1, String Filtervalue) throws IOException {
    	
    	//
	
		//int Sum=0;

    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
       sheet.setAutoFilter(CellRangeAddress.valueOf("A1:I1")); 
       
       /* Step-1: Get the CTAutoFilter Object */
       CTAutoFilter sheetFilter=sheet.getCTWorksheet().getAutoFilter();
       
       
       /* Step -2: Add new Filter Column */
       CTFilterColumn  myFilterColumn=sheetFilter.insertNewFilterColumn(0);
       
       
      
       
       /* Step-3: Set Filter Column ID */
       myFilterColumn.setColId(colId1);
       
       /* Step-4: Add new Filter */
       CTFilter myFilter=myFilterColumn.addNewFilters().insertNewFilter(0);
       /* Step -5: Define Auto Filter Condition - We filter Brand with Value of "A" */
       myFilter.setVal(Filtervalue);                           
       XSSFRow r1;
       
       int Sum=0;
       
       /* Step-6: Loop through Rows and Apply Filter */
       for(Row r : sheet) {
               for (Cell c : r) {
            	   
            	   
                 //  Sum  = Sum + (int)cell.getNumericCellValue(); 
                   
                   //System.out.println("Each Cell value is  : "+(int)cell.getNumericCellValue());  
                   
            	   
                       if (c.getColumnIndex()== columnIndex && !c.getStringCellValue().equals(Filtervalue)) {
                               r1=(XSSFRow) c.getRow();
                               if (r1.getRowNum()!=0) { /* Ignore top row */
                                       /* Hide Row that does not meet Filter Criteria */
                                       r1.getCTRow().setHidden(true); 
                                    
                               				}
                                       }       
              
               
               
               }
               
            
               

                       
           /*    for(Row r11 : sheet) {
                   for (Cell c : r11) { 
                	   
                if (c.getColumnIndex()== columnIndex && !c.getStringCellValue().equals(Filtervalue)) {
                    // if(cell.getCellTypeEnum() == CellType.NUMERIC){	
                    r1=(XSSFRow) c.getRow();
                    
		                    if (r1.getRowNum()!=0) { 
		                       	 Sum  = Sum + (int)cell.getNumericCellValue(); 
		                    }
                       	 
          
                       }
                   }
                   
                   System.out.println("Total Sum is : "+Sum);
               }*/
                       
                     /*  if (c.getColumnIndex()== columnIndex && !c.getStringCellValue().equals(Filtervalue)) {
                           r1=(XSSFRow) c.getRow();
                           if (r1.getRowNum()!=0) { /* Ignore top row 
                               //  Hide Row that does not meet Filter Criteria 
                                   r1.getCTRow().setHidden(true); }
                                   }   
                       
                       
               }*/
       }
       
       

       
       fout= new FileOutputStream(xlFilePath);
   		workbook.write(fout);
   		
   		
      
       fout.flush();
       fout.close();
        workbook.close();
        fis.close();

        
    }
    
    
    
    
    


public  void PrtintAll_Data(String xlFilePath,String sheetName, int columnIndex,long colId1, String Filtervalue) throws IOException {
   	
   	//http://thinktibits.blogspot.com/2014/09/Excel-auto-filter-criteria-Apache-POI-Example.html
	
		//int Sum=0;

   	fis = new FileInputStream(xlFilePath);
       workbook = new XSSFWorkbook(fis);
       sheet = workbook.getSheet(sheetName);
     // sheet.setAutoFilter(CellRangeAddress.valueOf("A1:I1")); 
      
                        
    //  XSSFRow r11 = null;
      
      int Sum=0;
      
    
      // Creating formula evaluator object
      FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
      
      Iterator iterator = sheet.iterator();
     
      while (iterator.hasNext()) {
       Row nextRow = (Row) iterator.next();
       Iterator cellIterator = nextRow.cellIterator();
     
       while (cellIterator.hasNext()) {
        Cell cell = (Cell) cellIterator.next();
     
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_STRING:
         System.out.print(cell.getStringCellValue());
         break;
        case Cell.CELL_TYPE_NUMERIC:
         System.out.print(cell.getNumericCellValue());
         break;
        case Cell.CELL_TYPE_BOOLEAN:
         System.out.print(cell.getBooleanCellValue());
         break;
        case Cell.CELL_TYPE_FORMULA:
        {
         // Evaluting cell
          CellValue c=formulaEval.evaluate(cell);
         System.out.println(c.getNumberValue());
        }
        }
        System.out.print(" | ");
       }
       System.out.println();
      }
      

      
     fout= new FileOutputStream(xlFilePath);
  		workbook.write(fout);
  		
  		
     
      fout.flush();
      fout.close();
       workbook.close();
       fis.close();

       
   }
   
   
    
    
    
    


public  void setAutoFilter_Single1_Sum1(String xlFilePath,String sheetName, int columnIndex,long colId1, String Filtervalue) throws IOException {
   	
   	//http://thinktibits.blogspot.com/2014/09/Excel-auto-filter-criteria-Apache-POI-Example.html
	
		//int Sum=0;

   	fis = new FileInputStream(xlFilePath);
       workbook = new XSSFWorkbook(fis);
       sheet = workbook.getSheet(sheetName);
     // sheet.setAutoFilter(CellRangeAddress.valueOf("A1:I1")); 
 		sheet.isColumnHidden(3);
                        
    //  XSSFRow r11 = null;
      
      int Sum=0;
   

      
     //fout= new FileOutputStream(xlFilePath);
  	//	workbook.write(fout);

  		
     
      //fout.flush();
    //  fout.close();
       workbook.close();
       fis.close();

       
   }
   
   
    
    
    
    
    
    
    
    
    
    public  void setAutoFilter_Single(String xlFilePath,String sheetName, int column, String value) throws IOException {
    	
    	//http://thinktibits.blogspot.com/2014/09/Excel-auto-filter-criteria-Apache-POI-Example.html

    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
       sheet.setAutoFilter(CellRangeAddress.valueOf("A1:G1")); 
       
       /* Step-1: Get the CTAutoFilter Object */
       CTAutoFilter sheetFilter=sheet.getCTWorksheet().getAutoFilter();                             
       /* Step -2: Add new Filter Column */
       CTFilterColumn  myFilterColumn=sheetFilter.insertNewFilterColumn(0);
       /* Step-3: Set Filter Column ID */
       myFilterColumn.setColId(1L);
       /* Step-4: Add new Filter */
       CTFilter myFilter=myFilterColumn.addNewFilters().insertNewFilter(0);
       /* Step -5: Define Auto Filter Condition - We filter Brand with Value of "A" */
       myFilter.setVal("A");                           
       XSSFRow r1;
       /* Step-6: Loop through Rows and Apply Filter */
       for(Row r : sheet) {
               for (Cell c : r) {
                       if (c.getColumnIndex()==1 && !c.getStringCellValue().equals("A")) {
                               r1=(XSSFRow) c.getRow();
                               if (r1.getRowNum()!=0) { /* Ignore top row */
                                       /* Hide Row that does not meet Filter Criteria */
                                       r1.getCTRow().setHidden(true); }
                                       }                               
               }
       }
       
       

       
       fout= new FileOutputStream(xlFilePath);
   		workbook.write(fout);
   		
   		
      
       fout.flush();
       fout.close();
        workbook.close();
        fis.close();

        
    }
    

    
    public  void setAutoFilter_Multiple(String xlFilePath,String sheetName, int column, String value) throws IOException {
    	
    	//http://thinktibits.blogspot.com/2014/09/Excel-auto-filter-criteria-Apache-POI-Example.html

    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
       sheet.setAutoFilter(CellRangeAddress.valueOf("A1:G1")); 
       
       /* Step-1: Get the CTAutoFilter Object */
       CTAutoFilter sheetFilter=sheet.getCTWorksheet().getAutoFilter();                             
       /* Step -2: Add new Filter Column */
       CTFilterColumn  myFilterColumn=sheetFilter.insertNewFilterColumn(0);
       /* Step-3: Set Filter Column ID */
       myFilterColumn.setColId(1L);
       /* Step-4: Add Multiple Filters on a Single Column */
       CTFilters listofFilters=myFilterColumn.addNewFilters();
       CTFilter myFilter1=listofFilters.addNewFilter();
       CTFilter myFilter2=listofFilters.addNewFilter();
       /* Step-5: Define Multiple Filters */
       myFilter1.setVal("A");  // Filter by Both A and B
       myFilter2.setVal("B");                  
       /* Add this to a list for comparison */
       List<String> list = new ArrayList<String>();
       list.add("A");
       list.add("B");          
       XSSFRow r1;
       /* Step-6: Loop through Rows and Apply Filter */
       for(Row r : sheet) {
               for (Cell c : r) {
                       if (c.getColumnIndex()==1 && !list.contains(c.getStringCellValue())) {
                               r1=(XSSFRow) c.getRow();
                               if (r1.getRowNum()!=0) { /* Ignore top row */
                                       /* Hide Row that does not meet Filter Criteria */
                                       r1.getCTRow().setHidden(true); }
                       }                               
               }
       }
       
       

       
       fout= new FileOutputStream(xlFilePath);
   		workbook.write(fout);
   		
   		
      
       fout.flush();
       fout.close();
        workbook.close();
        fis.close();

        
    }
    
    
    
    
    
    
    
    
    
    
    

    public  void setAutoFilter_Multiple_Columns(String xlFilePath,String sheetName, int column, String value) throws IOException {
    	
    	//http://thinktibits.blogspot.com/2014/09/Excel-auto-filter-criteria-Apache-POI-Example.html

    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
       sheet.setAutoFilter(CellRangeAddress.valueOf("A1:G1")); 
   
       
       /* Step-1: Get the CTAutoFilter Object */
       CTAutoFilter sheetFilter=sheet.getCTWorksheet().getAutoFilter();             

       /* Step -2: Filter Multiple Columns in POI with Auto Filter */
       CTFilterColumn  myFilterColumn=sheetFilter.insertNewFilterColumn(0);            
       myFilterColumn.setColId(1L);

       CTFilterColumn  myFilterColumn2=sheetFilter.insertNewFilterColumn(1);           
       myFilterColumn2.setColId(5L);
                       
       CTFilters firstColumnFilter=myFilterColumn.addNewFilters();
       CTFilter myFilter1=firstColumnFilter.addNewFilter();

       CTFilters secondColumnFilter=myFilterColumn2.addNewFilters();
       CTFilter myFilter2=secondColumnFilter.addNewFilter();

       /* Define Filter Conditions Across both columns */

       myFilter1.setVal("A");
       myFilter2.setVal("14512");   
              

       /* Add Filter Conditions to List */

       List<String> list1 = new ArrayList<String>();
       list1.add("A");

       List<String> list2 = new ArrayList<String>();
       list2.add("14512");

       /* Refresh Records to Match Multiple Filter Conditions */
       XSSFRow r1;
       /* Step-6: Loop through Rows and Apply Filter */
       for(Row r : sheet) {
               for (Cell c : r) {
                       //big IF condition to check both columns with value
                       if ( (c.getColumnIndex()==0 && !list2.contains(c.getStringCellValue())) || (c.getColumnIndex()==1 && !list1.contains(c.getStringCellValue())) ){
                               r1=(XSSFRow) c.getRow();
                               if (r1.getRowNum()!=0) { /* Ignore top row */
                               /* Hide Row that does not meet Filter Criteria */
                                       r1.getCTRow().setHidden(true); }
                       }                               
               }
       }

       
       fout= new FileOutputStream(xlFilePath);
   		workbook.write(fout);
   		
   		
      
       fout.flush();
       fout.close();
        workbook.close();
        fis.close();

        
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    public void clearsheetdata(String xlFilePath, String sheetName) throws Exception
    {
    	

    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
        
        for(int i=workbook.getNumberOfSheets()-1;i>=0;i--){
            XSSFSheet tmpSheet =workbook.getSheetAt(i);
            if(tmpSheet.getSheetName().equals(sheetName)){
            	workbook.removeSheetAt(i);
            }
        }       

        
        fout= new FileOutputStream(xlFilePath);
    	workbook.write(fout);
     
        fout.flush();
        fout.close();
        workbook.close();
        fis.close();
        
        
    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        workbook.createSheet(sheetName);
        
        fout= new FileOutputStream(xlFilePath);
    	workbook.write(fout);
     
        fout.flush();
        fout.close();
        workbook.close();
        fis.close();
        
        
        
    }
    

    
    public void CopyWorkBook(String xlFilePathcurr,String xlFilePathnew) throws Exception
    {
    	//input source excel file which contains sheets to be copied
    	System.out.println("started : ");
    	fis = new FileInputStream(xlFilePathcurr);
    	workbook = new XSSFWorkbook(fis);
    	
    	XSSFWorkbook workbookoutput=workbook;
    	
    	//fout= new FileOutputStream(xlFilePath);  
    	fout= new FileOutputStream(xlFilePathnew);
    	workbookoutput.write(fout);

    	//output new excel file to which we need to copy the above sheets
    	//this would copy entire workbook from source
    	//XSSFWorkbook workbookoutput=workbook;
    	

        fout.flush();
        fout.close();
        workbook.close();
        fis.close();
        
    	System.out.println("completed : ");
    }
    
    
  //getRow count starts
 
    public int getRowCount(String xlFilePath, String sheetName) throws Exception
    {
    	
        fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(0);
       
       
       // int colNum = row.getLastCellNum();
      //  System.out.println("Total Number of Columns in the excel is : "+colNum);
        int RowNum = sheet.getLastRowNum()+1;
        System.out.println("Total Number of Rows in the excel is : "+RowNum);
        
        workbook.close();
    	fis.close();
        
        return RowNum;
        
    
    }
    
    //getRow count ends
    
    
    
    //getRow count starts
   
      public int getColumnCount(String xlFilePath, String sheetName) throws Exception
      {
      	
          fis = new FileInputStream(xlFilePath);
          workbook = new XSSFWorkbook(fis);
          sheet = workbook.getSheet(sheetName);
          row = sheet.getRow(0);
         
         
          int colNum = row.getLastCellNum();
         System.out.println("Total Number of Columns in the excel is : "+colNum);
        // int RowNum = sheet.getLastRowNum()+1;
        //  System.out.println("Total Number of Rows in the excel is : "+RowNum);
          
          workbook.close();
      	fis.close();
          
          return colNum;
          
      
      }
      
      //getColumn count ends
    
    
 
    public String getCellData(String xlFilePath, String sheetName,int rowNum,int column) throws Exception
    {
    	
    	
    	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(rowNum);
        cell = row.getCell(column);
        
        FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
        ((XSSFFormulaEvaluator) formulaEval).evaluateAllFormulaCells(workbook);
        
       // XSSFFormulaEvaluator.evaluateAllFormulaCells(xssfWorkbook)
        
        if (cell.getCellTypeEnum() == CellType.FORMULA)
        {
        	
        	//String str6="NoData";
            // Creating formula evaluator object
            
        	
            System.out.println("Formula is " + cell.getCellFormula());
           System.out.println("Last evaluated as: " + cell.getNumericCellValue());
                

            
            
        	//  CellValue c=formulaEval.evaluate(cell);
             // System.out.println(c.getNumberValue());
              
              
              
              
              
              
          	int i = (int)cell.getNumericCellValue(); 
        	 String str6 = String.valueOf(i); 
        
        	//String str6=cell.getStringCellValue();
        	//workbook.close();
            
              workbook.close();
              fis.close();
            return str6;
        }
               
        else if(cell.getCellTypeEnum() == CellType.STRING)
        {	
        	String str6=cell.getStringCellValue();
        	workbook.close();
        	fis.close();
            return str6;
        }
        else if (cell.getCellTypeEnum() == CellType.NUMERIC)
        {	
        	//int str6=cell.getStringCellValue();
        	int i = (int)cell.getNumericCellValue(); 
        	String str6 = String.valueOf(i); 
        	workbook.close();
        	fis.close();
            return str6;
        }
        else
        {
        	String str6=cell.getStringCellValue();
        	workbook.close();
        	fis.close();
            return str6;
        }
		
        
        
        
        
        
    } 
    
    
    
    public  void PutCellData(String xlFilePath,String sheetName,int rowNum,int column,String Text)
    		throws Exception
    {
    	FileInputStream fis = null;
	   	XSSFWorkbook workbook = null;
	   	XSSFSheet sheet = null;
	   	XSSFRow row = null;
	    XSSFCell cell = null;
	   	FileOutputStream fout=null;
   	 
   	 	fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
    	sheet = workbook.getSheet(sheetName);
    	
    	if(sheet.getRow(rowNum)==null)
    	{
    		row=sheet.createRow(rowNum);
    	}
    	else
    	{
    		row=sheet.getRow(rowNum);
    	}
    	
    	
    	if(row.getCell(column)==null)
    	{
    		cell=row.createCell(column);
    	}
    	else
    	{
    		cell=row.getCell(column);
    	}
    	
    	//cell.getSheet();
    	//Cell cell = null; // declare a Cell object
    	
    	// Access the second cell in second row to update the value
    	
    	
    	cell = sheet.getRow(rowNum).getCell(column);  
    	cell.setCellValue(Text);
    	
    	/*
    	CellStyle backgroundStyle = workbook.createCellStyle();
        backgroundStyle.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        backgroundStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
         cell.setCellStyle(backgroundStyle);
      */
         
         CellStyle cs1 = workbook.createCellStyle(); 
         cs1.setFillForegroundColor(IndexedColors.WHITE.getIndex()); 
         cs1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
     
         Font font = workbook.createFont();
         font.setColor(IndexedColors.BLUE.getIndex());
         font.setBold(false);
         cs1.setFont(font);
   
    	
    	System.out.println("Text:"+Text);
    	cell.setCellStyle(cs1);
    	
    	cell.setCellValue(Text);
    	
    	

    	//fout= new FileOutputStream(xlFilePath); 
    	
    	fout= new FileOutputStream(xlFilePath);
    	workbook.write(fout);
     
        fout.flush();
        fout.close();
        workbook.close();
        fis.close();
        
    
       /*
        * https://stackoverflow.com/questions/41057500/not-able-to-update-cell-value-using-setcellvalue-api-in-java
       
        fis= new FileInputStream(new File(xlFilePath)); 
       workbook = new XSSFWorkbook(fis);; //Access the workbook
       sheet = workbook.getSheet(sheetName); //Access the worksheet, so that we can update / modify it.
        		// cell = null; // declare a Cell object
       
       row=sheet.createRow(rowNum);
   	cell=row.createCell(column);
        		cell = sheet.getRow(rowNum).getCell(column);   // Access the second cell in second row to update the value
        		cell.setCellValue(Text);  // Get current cell value value and overwrite the value
        		fis.close(); //Close the InputStream
        		FileOutputStream output_file =new FileOutputStream(new File(xlFilePath));  
        		workbook.write(output_file); //write changes
        		output_file.close();  //close the stream
        
         */
        
    }
    
    
    
    public  void PutCellFormula(String xlFilePath,String sheetName,int rowNum,int column,String Text)throws Exception
    {
    	
        
     
    	
       // System.out.println("FilePath : "+xlFilePath);
     	//System.out.println("sheetName : "+sheetName);
     	//System.out.println("colNum : "+column);
     	//System.out.println("rowNum : "+rowNum);
        // cell =sheet.createRow(rowNum).createCell(column);
       // cell.setHyperlink(link);
       // cell.setCellStyle(hlink_style);
        //cell.setCellValue(Text1);
        

        FileInputStream fis = null;
       	 XSSFWorkbook workbook = null;
       	 XSSFSheet sheet = null;
       	XSSFRow row = null;
        XSSFCell cell = null;
       	 FileOutputStream fout=null;
       	 
     fis = new FileInputStream(xlFilePath);
            workbook = new XSSFWorkbook(fis);
        	sheet = workbook.getSheet(sheetName);
        	
        	CreationHelper createHelper = workbook.getCreationHelper();
            CellStyle hlink_style = workbook.createCellStyle();
            Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(Text);
        	
        	if(sheet.getRow(rowNum)==null)
        	{
        	row=sheet.createRow(rowNum);
        	}
        	else
        	{
        		row=sheet.getRow(rowNum);
        	}
        	
        	
        	
        	
        	if(row.getCell(column)==null)
        	{
        	cell=row.createCell(column);
        	}
        	else
        	{
        		cell=row.getCell(column);
        	}
        	
        	//cell.getSheet();
        	//Cell cell = null; // declare a Cell object
        	cell = sheet.getRow(rowNum).getCell(column);   // Access the second cell in second row to update the value
        	cell.setCellValue(Text);
        	
        //	System.out.println("Text:"+Text);
        	
        	cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
        	cell.setCellValue(Text);

        	//fout= new FileOutputStream(xlFilePath);  
        	fout= new FileOutputStream(xlFilePath);
        	workbook.write(fout);
         
            fout.flush();
            fout.close();
            workbook.close();
            fis.close();
         

    }
    
    
/*
    
    try {
        ExcelApiTest3 eat=new ExcelApiTest3();
 		eat.PutCellData("E://TC01.xls","Sheet1",0,9, "Actual Result");	
    }catch(Exception e) {}
    
    */
    
    
    public  void PassPutCellData(String xlFilePath,String sheetName,int rowNum,int column,String Text)throws Exception
    {
    	
    FileInputStream fis = null;
   	 XSSFWorkbook workbook = null;
   	 XSSFSheet sheet = null;
   	XSSFRow row = null;
    XSSFCell cell = null;
   	 FileOutputStream fout=null;
   	 
 fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
    	sheet = workbook.getSheet(sheetName);
    	
    	if(sheet.getRow(rowNum)==null)
    	{
    	row=sheet.createRow(rowNum);
    	}
    	else
    	{
    		row=sheet.getRow(rowNum);
    	}
    	
    	
    	
    	
    	if(row.getCell(column)==null)
    	{
    	cell=row.createCell(column);
    	}
    	else
    	{
    		cell=row.getCell(column);
    	}
    	
    	//cell.getSheet();
    	//Cell cell = null; // declare a Cell object
    	cell = sheet.getRow(rowNum).getCell(column);   // Access the second cell in second row to update the value
    	cell.setCellValue(Text);
    	/*
    	CellStyle backgroundStyle = workbook.createCellStyle();
        backgroundStyle.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        backgroundStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
         cell.setCellStyle(backgroundStyle);
      */
         
         CellStyle cs1 = workbook.createCellStyle(); 
         cs1.setFillForegroundColor(IndexedColors.GREEN.getIndex()); 
         cs1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         //cs1.setFillPattern(CellStyle.SOLID_FOREGROUND); 
         
         Font font = workbook.createFont();
         font.setColor(IndexedColors.WHITE.getIndex());
         font.setBold(true);
         cs1.setFont(font);
   
    	
    	System.out.println("Text:"+Text);
    	cell.setCellStyle(cs1);
    	cell.setCellValue(Text);

    	//fout= new FileOutputStream(xlFilePath);  
    	fout= new FileOutputStream(xlFilePath);
    	workbook.write(fout);
     
        fout.flush();
        fout.close();
        workbook.close();
        fis.close();
        
    
       /*
        * https://stackoverflow.com/questions/41057500/not-able-to-update-cell-value-using-setcellvalue-api-in-java
       
        fis= new FileInputStream(new File(xlFilePath)); 
       workbook = new XSSFWorkbook(fis);; //Access the workbook
       sheet = workbook.getSheet(sheetName); //Access the worksheet, so that we can update / modify it.
        		// cell = null; // declare a Cell object
       
       row=sheet.createRow(rowNum);
   	cell=row.createCell(column);
        		cell = sheet.getRow(rowNum).getCell(column);   // Access the second cell in second row to update the value
        		cell.setCellValue(Text);  // Get current cell value value and overwrite the value
        		fis.close(); //Close the InputStream
        		FileOutputStream output_file =new FileOutputStream(new File(xlFilePath));  
        		workbook.write(output_file); //write changes
        		output_file.close();  //close the stream
        
         */
        
    } 
    
    
    ////////////////////
    

    public  void FailPutCellData(String xlFilePath,String sheetName,int rowNum,int column,String Text)throws Exception
    {
    	
    FileInputStream fis = null;
   	 XSSFWorkbook workbook = null;
   	 XSSFSheet sheet = null;
   	XSSFRow row = null;
    XSSFCell cell = null;
   	 FileOutputStream fout=null;
   	 
 fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
    	sheet = workbook.getSheet(sheetName);
    	
    	if(sheet.getRow(rowNum)==null)
    	{
    	row=sheet.createRow(rowNum);
    	}
    	else
    	{
    		row=sheet.getRow(rowNum);
    	}
    	
    	
    	
    	
    	if(row.getCell(column)==null)
    	{
    	cell=row.createCell(column);
    	}
    	else
    	{
    		cell=row.getCell(column);
    	}
    	
    	//cell.getSheet();
    	//Cell cell = null; // declare a Cell object
    	cell = sheet.getRow(rowNum).getCell(column);   // Access the second cell in second row to update the value
    	cell.setCellValue(Text);
    	/*
    	CellStyle backgroundStyle = workbook.createCellStyle();
        backgroundStyle.setFillBackgroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        backgroundStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
         cell.setCellStyle(backgroundStyle);
      */
         
         CellStyle cs1 = workbook.createCellStyle(); 
         cs1.setFillForegroundColor(IndexedColors.RED.getIndex()); 
         cs1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
         //cs1.setFillPattern(CellStyle.SOLID_FOREGROUND); 
         
         Font font = workbook.createFont();
         font.setColor(IndexedColors.WHITE.getIndex());
         font.setBold(true);
         cs1.setFont(font);
   
    	
    	System.out.println("Text:"+Text);
    	cell.setCellStyle(cs1);
    	cell.setCellValue(Text);

    	//fout= new FileOutputStream(xlFilePath);  
    	fout= new FileOutputStream(xlFilePath);
    	workbook.write(fout);
     
        fout.flush();
        fout.close();
        workbook.close();
        fis.close();
        
    
       /*
        * https://stackoverflow.com/questions/41057500/not-able-to-update-cell-value-using-setcellvalue-api-in-java
       
        fis= new FileInputStream(new File(xlFilePath)); 
       workbook = new XSSFWorkbook(fis);; //Access the workbook
       sheet = workbook.getSheet(sheetName); //Access the worksheet, so that we can update / modify it.
        		// cell = null; // declare a Cell object
       
       row=sheet.createRow(rowNum);
   	cell=row.createCell(column);
        		cell = sheet.getRow(rowNum).getCell(column);   // Access the second cell in second row to update the value
        		cell.setCellValue(Text);  // Get current cell value value and overwrite the value
        		fis.close(); //Close the InputStream
        		FileOutputStream output_file =new FileOutputStream(new File(xlFilePath));  
        		workbook.write(output_file); //write changes
        		output_file.close();  //close the stream
        
         */
        
    }
    
    
    //////////////////////
    

    

    ////////////////////// Data Provider starts///////////////////////////////////////
    
    
    public  Object[][] getTableArray(String FilePath, String SheetName) throws Exception {   
    	 
        
    	//System.out.println("Hai I am here");
    //	System.out.println("FilePath :"+FilePath);
    //	System.out.println("SheetName :"+SheetName);
    	
    	String[][] tabArray = null;
        
    
     
        try { 
     
        FileInputStream fis = new FileInputStream(FilePath);
     
        // Access the required test data sheet
     
        //ExcelWBook = new HSSFWorkbook(ExcelFile);
        workbook = new XSSFWorkbook(fis);
     
       // ExcelWSheet = ExcelWBook.getSheet(SheetName);
        sheet = workbook.getSheet(SheetName);
        
       // System.out.println("FilePath3 :"+FilePath);
    	//System.out.println("SheetNam3 :"+SheetName);
    	//System.out.println("Sheet 1 ,1 value :" +getCellData(1,1));
     
        int startRow = 1;
     
        int startCol = 0;
     
        int ci,cj;
     
        int totalRows = sheet.getLastRowNum();
     
        // you can write a function as well to get Column count
     
       // int totalCols = 2;
        
        int totalCols = getColumnCount(FilePath,SheetName);
        //totalCols=totalCols-1;
       // totalRows=totalRows-1;
        
        
        System.out.println("totalRows0 :"+totalRows);
        System.out.println("totalCols0 :"+totalCols);
     
        tabArray=new String[totalRows][totalCols];
     
     //   System.out.println("totalRows1 :"+totalRows);
      //  System.out.println("totalCols1 :"+totalCols);

      
        ci=0;
        
        for (int i=startRow;i<=totalRows;i++, ci++) {              
     
       cj=0;
     
        for (int j=startCol;j<totalCols;j++, cj++){
     
        tabArray[ci][cj]=getCellData(i,j);
     
       // System.out.println(tabArray[ci][cj]);  
     
     }
     
     }
     
        
        
     }
     
     catch (FileNotFoundException e){
     
     System.out.println("File not Found Exception");
     
     e.printStackTrace();
     
     }
     
     catch (IOException e){
     
     System.out.println("Could not read the Excel sheet");
     
     e.printStackTrace();
     
     }
     
     return(tabArray);
     
     }
     
public  String getCellData(int RowNum, int ColNum) throws Exception {
     
     try{
     
	    cell = sheet.getRow(RowNum).getCell(ColNum);
		 // String CellData = cell.getStringCellValue();
		     //return CellData;
		 int dataType = cell.getCellType();
		     
		     if  (dataType == 3) {
		     
		     return "";
		     
		     }else{
		     
		     String CellData = cell.getStringCellValue();
		     
		     return CellData;
		     
		     }
	
		     
		    
		     
     }catch (Exception e){
     
  System.out.println(e.getMessage());
     
 throw (e);
     
     }
  }
 
 /*
public static String getCellData(int rowNum,int column) throws Exception
    {
    	//fis = new FileInputStream(xlFilePath);
      //  workbook = new HSSFWorkbook(fis);
      //  sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(rowNum);
        cell = row.getCell(column);
       
        if(cell.getCellTypeEnum() == CellType.STRING)
        {	
        	String str6=cell.getStringCellValue();
        	workbook.close();
        	fis.close();
            return str6;
        }
        else 
        {
        	String str6=cell.getStringCellValue();
        	workbook.close();
        	fis.close();
            return str6;
        }
        
        
        
        
    } */
    
 
     //////////////////////////////////////Data Provider ends//////////////////////
    
    
}

