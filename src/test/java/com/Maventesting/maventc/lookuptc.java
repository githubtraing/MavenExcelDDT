package com.Maventesting.maventc;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class lookuptc {
	WebDriver driver;
	String URL="https://www.aritzia.com/";
	String strxlFilepath="D:\\Eclipse_workspace_april\\TestData\\SignUpAccount.xls";
	int xlNumOfRows;
	int xlNumOfCols;
	
	String xlDataInLocalArray[][];// declare xlDataInLocalArray[][] is an Array

	@Test
	public void tc_xlAccountSignUp() throws IOException {
		
		
		ReadDatafromExcel(strxlFilepath);
	 	/*
	 	String EmAddress= "jessica-cai11@hotmail.com";	 	 
		String Password = "abc123";
		String FirstName ="jessica";
		String LastName = "Cai";
		*/
		
		for (int i=1;i<xlNumOfRows;i++) {
			
			String EmAddress= xlDataInLocalArray[i][0];	 	 
			String Password =xlDataInLocalArray[i][1];	
			String FirstName =xlDataInLocalArray[i][2];	
			String LastName = xlDataInLocalArray[i][3];	
		System.setProperty("webdriver.chrome.driver", "D:\\ChromeDriver\\chromedriver_win32\\chromedriver.exe");		
	    driver =new ChromeDriver();
	    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	    driver.manage().window().maximize();
	    driver.get(URL);
		driver.findElement(By.xpath("//*[@id=\"loggedout\"]/a")).click();
		//driver.findElement(By.linkText("SIGN IN")).click();
		driver.findElement(By.linkText("Create an Account")).click();
		driver.findElement(By.id("dwfrm_profile_customer_firstname")).sendKeys(FirstName);
		driver.findElement(By.id("dwfrm_profile_customer_lastname")).sendKeys(LastName);
		driver.findElement(By.id("dwfrm_profile_customer_email")).sendKeys(EmAddress);
		driver.findElement(By.id("dwfrm_profile_login_password")).sendKeys(Password);
		driver.quit();
		}
	}
		public void ReadDatafromExcel(String strxlFilepath) throws IOException {
			//we have a xlfile which locate at strxlFilepath
			File xlFile= new File(strxlFilepath);
			//to make read file faster which need to use FileInputStream,use this to point to the xlFile and read this file
		    FileInputStream TestDataStream = new FileInputStream(xlFile);

 
HSSFWorkbook xlWorkbook= new HSSFWorkbook(TestDataStream); // after point to the file, will tell to read from the workbook
            
            HSSFSheet xlsheet=xlWorkbook.getSheetAt(0);  //reffering to 1st sheet


            xlNumOfRows= xlsheet.getLastRowNum()+1;           
            xlNumOfCols= xlsheet.getRow(0).getLastCellNum();
            		
System.out.println(".....");
            System.out.println("Total Number of Test-data Rows are"+xlNumOfRows);
            System.out.println("Total Number of Test-data Cols are"+xlNumOfCols);   
            
            xlDataInLocalArray = new String[xlNumOfRows][xlNumOfCols];
            
    // fill these data from excel into Array        
   for(int i=0; i<xlNumOfRows;i++)   {
        	 HSSFRow row = xlsheet.getRow(i);
        	 for(int j=0; j< xlNumOfCols;j++) {
        		 HSSFCell cell=row.getCell(j);//To read value from each col in each row

             String value= cellToString(cell);
             xlDataInLocalArray[i][j]=value;
            // System.out.print(value);  
            // Syetem.out.print("@@");             
        	 }
        System.out.println();	 
         }

} 
	
public static String cellToString(HSSFCell cell)	{
	    //This function will convert an object of type excel cell to a string value
       int type=cell.getCellType();
       Object result = null;
       switch (type) {
       case  HSSFCell.CELL_TYPE_NUMERIC://0
    	   result = cell.getNumericCellValue();
    	   break;
       case HSSFCell.CELL_TYPE_STRING://1
    	   result= cell.getStringCellValue();
    	   break;
       case HSSFCell.CELL_TYPE_FORMULA://2
    	   throw new RuntimeException("we can't evalute formulas is Java");
       case HSSFCell.CELL_TYPE_BLANK://3
    	   result= "-";
    	   break;
       case HSSFCell.CELL_TYPE_BOOLEAN://4
    	   result= cell.getBooleanCellValue();
    	   break;
       case HSSFCell.CELL_TYPE_ERROR://5
    	   throw new RuntimeException("this cell has an error");
    	   
    	   
       }
       return result.toString();
}
		
		@Test
		public void tc_searchitems() {
			System.out.println("no code to run in here...serchitems");
		}
	
		@Test
		public void tc_searchnew() {
			System.out.println("no code to run in here...serchnew");
		}
	

	
	
}
