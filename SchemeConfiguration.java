package com.innoviti.selenium;


import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.DataProvider;


public class SchemeConfiguration{

	WebDriver driver;
	
	@Test(dataProvider = "data1")
	public void script(String[] inp)
	{	
		try {
			
			System.setProperty("webdriver.chrome.driver","C:\\Selenium\\chromedriver.exe");
			WebDriver driver = new ChromeDriver();
			
			System.out.println("Starting Test...");
			
			JavascriptExecutor js = (JavascriptExecutor) driver;
	        driver.manage().deleteAllCookies();
	        driver.manage().window().maximize();
	        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
	        driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
	        
	        /*LOGIN */
	        
	        driver.get("https://192.168.0.83:20415/portal");
	        WebElement user = driver.findElement(By.id("j_username"));
	        WebElement password = driver.findElement(By.id("j_password"));
	        user.sendKeys("admin");
	        password.sendKeys("Innoviti@123");
	        
	        WebElement login = driver.findElement(By.xpath("//*[@id=\"loginForm\"]/div[4]/button"));
	        login.click();
	        
	        /*SCHEME CONFIGURATION*/
	        
	        WebElement emiconfig = driver.findElement(By.xpath("/html/body/div[1]/div/div[1]/div[1]/div[2]/div/ul/li[4]"));
			emiconfig.click();
			
			WebElement admin = driver.findElement(By.xpath("//*[@id=\"treeContent\"]/ul/li[1]"));
			admin.click();
			
			//Thread.sleep(500);
			
			WebElement schconfig = driver.findElement(By.xpath("//*[@id=\"schemeConfigurationTab\"]/a"));
	        schconfig.click();
	        
	        
	        //Adding new scheme
	        WebElement new_sch = driver.findElement(By.xpath("//*[@id=\"schemeConfigurationMainContent\"]/div[1]/span/a"));
	        new_sch.click();
	        
	        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
	        
	        WebElement sch_name = driver.findElement(By.xpath("//*[@id=\"emiSchemeName\"]"));
	        sch_name.sendKeys(inp[0]);
	        
	        Select bank = new Select(driver.findElement(By.id("emiBank")));
		    bank.selectByVisibleText(inp[1]);
		    
		    
		    String date1 = inp[2];
		    String date2 = inp[3];
		    String[] sdate = date1.split("/");
		    String[] edate = date2.split("/");
		    System.out.println("SMonth:"+Integer.parseInt(sdate[1]));
		    
		    //Start_Date
		    WebElement start_date = driver.findElement(By.xpath("//*[@id=\"emiSchemeStartDate\"]"));
		    start_date.click();
		    
		    int smonth = Integer.parseInt(sdate[1]);
		    switch (smonth) {
            case 1:  sdate[1] = "Jan";       break;
            case 2:  sdate[1] = "Feb";      break;
            case 3:  sdate[1] = "Mar";         break;
            case 4:  sdate[1] = "Apr";         break;
            case 5:  sdate[1] = "May";           break;
            case 6:  sdate[1] = "Jun";          break;
            case 7:  sdate[1] = "Jul";          break;
            case 8:  sdate[1] = "Aug";        break;
            case 9:  sdate[1] = "Sep";     break;
            case 10: sdate[1] = "Oct";       break;
            case 11: sdate[1] = "Nov";      break;
            case 12: sdate[1] = "Dec";      break;
            default: sdate[1] = "Invalid"; 
        }
        System.out.println(sdate[1]);
		
        	int syear = Integer.parseInt(sdate[2]);
        	
        	boolean doubledigit = (syear > 9 && syear < 100);
        	
        	if(doubledigit)
        	{
        		Select year1 = new Select(driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]")));
            	year1.selectByVisibleText("20"+sdate[2]);
        	}
        	else
        	{
        		Select year1 = new Select(driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]")));
        		year1.selectByVisibleText(sdate[2]);
        	}
        
		    Select month1 = new Select(driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[1]")));
		    month1.selectByVisibleText(sdate[1]);
		    
		    WebElement day1 = driver.findElement(By.linkText(sdate[0]));
		    day1.click();
		    
		    js.executeScript("window.scrollBy(0,400)");
		    Thread.sleep(1000);
		    
		    //End_Date
		    WebElement end_date = driver.findElement(By.xpath("//*[@id=\"emiSchemeEndDate\"]"));
		    end_date.click();
		    System.out.println("EMonth:"+Integer.parseInt(edate[1]));
		    
		    int emonth = Integer.parseInt(edate[1]);
		    switch (emonth) {
            case 1:  edate[1] = "Jan";       break;
            case 2:  edate[1] = "Feb";      break;
            case 3:  edate[1] = "Mar";         break;
            case 4:  edate[1] = "Apr";         break;
            case 5:  edate[1] = "May";           break;
            case 6:  edate[1] = "Jun";          break;
            case 7:  edate[1] = "Jul";          break;
            case 8:  edate[1] = "Aug";        break;
            case 9:  edate[1] = "Sep";     break;
            case 10: edate[1] = "Oct";       break;
            case 11: edate[1] = "Nov";      break;
            case 12: edate[1] = "Dec";      break;
            default: edate[1] = "Invalid"; 
        }
        System.out.println(edate[1]);
		    
        	int eyear = Integer.parseInt(edate[2]);
    	
        	doubledigit = (eyear > 9 && eyear < 100);
        	System.out.println(edate[2]);
        	if(doubledigit)
        	{
        		Select year2 = new Select(driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]")));
            	year2.selectByVisibleText("20"+edate[2]);
        	}
        	else
        	{
        		Select year2 = new Select(driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]")));
            	year2.selectByVisibleText(edate[2]);
        	}
        
        	
        	
		    Select month2 = new Select(driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[1]")));
		    month2.selectByVisibleText(edate[1]);
		    
		    WebElement day2 = driver.findElement(By.linkText(edate[0]));
		    day2.click();
		    
		    WebElement min_amount = driver.findElement(By.xpath("//*[@id=\"emiSchemeMinAmount\"]"));
		    min_amount.sendKeys(inp[5]);
		    
		    WebElement max_amount = driver.findElement(By.xpath("//*[@id=\"emiSchemeMaxAmount\"]"));
		    max_amount.sendKeys(inp[4]);
		    
		    WebElement tenure = driver.findElement(By.xpath("//*[@id=\"emiSchemeTenure\"]"));
		    tenure.sendKeys(inp[6]);
		    
		    WebElement tenure_sel = driver.findElement(By.xpath("//*[@id=\"ui-id-1\"]"));
		    tenure_sel.click();
		    
		    Select advance_emi = new Select(driver.findElement(By.xpath("//*[@id=\"emiSchemeAdvanceEMI\"]")));
		    advance_emi.selectByVisibleText("3");
		    
		    WebElement interest = driver.findElement(By.xpath("//*[@id=\"emiSchemeInterestRate\"]"));
		    interest.sendKeys(inp[8]);
		    
		    //Processing fees
		    if(inp[9]=="F")
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeProcessingFeesTypeContainer\"]/label[1]")).click();
		    }
		    else
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeProcessingFeesTypeContainer\"]/label[2]")).click(); 
		    }
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeProcessingFees\"]")).sendKeys(inp[10]);
		    
		    
		    //Merchant Subvention 
		    
		    if(inp[13]=="U")
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeDealerSubvnContainer\"]/label[1]")).click();
		    }
		    else
		    {	
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeDealerSubvnContainer\"]/label[2]")).click(); 
		    }
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeDealerSubvnAmt\"]")).sendKeys(inp[14]);
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeDealerSubvnCapAmt\"]")).sendKeys(inp[15]);
		     
		    //Brand Subvention 
		    if(inp[17]=="U")
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeManufacturerSubvnContainer\"]/label[1]")).click();
		    else
		    driver.findElement(By.xpath("//*[@id=\"manufacturerSubvnCashbk\"]")).click();
		    
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeManufacturerSubvnAmt\"]")).sendKeys("300");
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeManufacturerSubvnCapAmt\"]")).sendKeys("800");
		    
		    //Bank Subvention 
		    if(inp[21]=="U")
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeBankSubvnContainer\"]/label[1]")).click();
		    else
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeBankSubvnContainer\"]/label[2]")).click(); 
		    
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeBankSubvnAmt\"]")).sendKeys("400");
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeBankSubvnCapAmt\"]")).sendKeys("900");
		    
		    
		    //Innoviti Subvention
		    if(inp[25]=="Ü")
		    driver.findElement(By.xpath("//*[@id=\"innovitiSubvnUpfrnt\"]")).click();
		    else
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeInnovitiSubvnContainer\"]/label[2]")).click(); 
		    
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeInnovitiSubvnAmt\"]")).sendKeys("200");
		    driver.findElement(By.xpath("//*[@id=\"emiSchemeInnovitiSubvnCapAmt\"]")).sendKeys("1000");
		    
		    js.executeScript("window.scrollBy(0,400)");
		    
		    //Cashback Type
		    if(inp[12]=="PRE")
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeCashbackTypeContainer\"]/label[2]")).click();
		    }
		    else
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeCashbackTypeContainer\"]/label[1]")).click();
		    }
		    //Subvention Criteria
		    if(inp[11]=="F")
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeSubventionCriteriaContainer\"]/label[1]")).click();
		    }
		    else
		    {
		    	driver.findElement(By.xpath("//*[@id=\"emiSchemeSubventionCriteriaContainer\"]/label[2]")).click();
		    }
		    	
		    
		    //Submit
		    driver.findElement(By.xpath("//*[@id=\"submitEMISchemeConfigBtn\"]")).click();
		    int t=1;
		    System.out.println("Test "+t+" Successfully Executed");
		    t++;
		    driver.close();
		    
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	


	

public static String[] DataDriven(int n) {
		String[] str=new String[29];
		try {
			// Specify the path of file
			File src=new File("C:\\Selenium\\EmiSchemes.xlsx");
			
			// load file
			FileInputStream fis=new FileInputStream(src);
			// Load workbook
			XSSFWorkbook wb=new XSSFWorkbook(fis);
			
			DataFormatter formatter = new DataFormatter();
			 

			XSSFSheet sh1= wb.getSheetAt(0);
			
			int colcount=sh1.getRow(0).getPhysicalNumberOfCells();
	      //  System.out.println("Total Number of Rows is ::"+rowcount);
	        System.out.println("Total number of Col is ::"+colcount);
	      //  String[] str=new String[colcount];
	        int k = n+1;
	        for(int i=n;i<k;i++){


	            for(int j=0;j<colcount;j++){

	            	 
	                String testdata1=formatter.formatCellValue(sh1.getRow(i).getCell(j));
	                System.out.println("Test data from excel cell"+j+"  :"+testdata1);
					str[j]= testdata1;
					
	                wb.close();
 

	            }

	        }
			
			
			
		
			
			
			
		} catch (FileNotFoundException e) {
			
			System.out.println(e.getMessage());
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
		return str;
		
		
	}


public static int rowcount(int r)
{
	try
	{
		File src=new File("C:\\Selenium\\EmiSchemes.xlsx");
	
		// load file
		FileInputStream fis=new FileInputStream(src);
		// Load workbook
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		
		XSSFSheet sh1= wb.getSheetAt(0);
		r = sh1.getLastRowNum();
		
	
	}
	catch(IOException e) {
		e.printStackTrace();
	
	}
	return r;
}

@DataProvider(name = "data1")
public void run()
{
	String[] inp = new String[29];
	int rc = rowcount(7);
	System.out.println("Total no. of rows: "+rc);
	for(int i=1;i<rc;i++)
	{
		inp=DataDriven(i);
		
		script(inp);
	}
}


public static void main(String[] args) 
{
	
	SchemeConfiguration obj=new SchemeConfiguration();
	obj.run();
	
	

	
	
}


}


