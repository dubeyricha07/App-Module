package appModules;
import java.io.FileInputStream;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.testng.Reporter;

import pageObjects.BaseClass;
import utility.Constant;
import utility.ExcelUtils;
import utility.Extend_Report;
import utility.Reusable;
     
   
    public class LenInfo {
    	
    	  private static XSSFSheet ExcelWSheet;
          private static XSSFWorkbook ExcelWBook;
          private static XSSFCell Cell;
          private static XSSFRow Row;
         
         
        public  static void Execute(int iTestCaseRow,String TC_Name) throws Exception{
        	WebDriver driver=BaseClass.driver;
        	String class_name1=Thread.currentThread().getStackTrace()[1].getClassName();
        	String[] parts1 = class_name1.split(Pattern.quote("."));
        	String class_name=parts1[1];
        	
        	try{      		
        	     
        		FileInputStream ExcelFile = new FileInputStream(Constant.Path_Excel + Constant.File_TestData);
            	//System.out.println(Constant.Path_Excel + Constant.File_TestData);
                // Access the required test data sheet
        		 
                ExcelWBook = new XSSFWorkbook(ExcelFile);
                ExcelWSheet = ExcelWBook.getSheet(class_name);
            	int noOfColumns = ExcelWSheet.getRow(0).getPhysicalNumberOfCells();                	
            	
				for (int i=1;i < noOfColumns;i++)
            	{
            	
					String  val=ExcelWSheet.getRow(0).getCell(i).getStringCellValue();
					
					int Valpos = val.indexOf("_");
					
					if (Valpos > 0)
					{
						
						String obj=val;
						String[] parts = val.split(Pattern.quote("_"));
						
						String Class_Name = parts[0];
						
						String Object_Type = parts[1];
						
						if (Class_Name.equalsIgnoreCase(class_name))
						{
							 
							

							switch (Object_Type){
							
							case "Edit":
								
								Reusable.Enter_Value_WebElement(driver,iTestCaseRow,obj, Class_Name,TC_Name);
								break;
							case "Button":
									
								Reusable.Click_Button(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								break;
							case "TypeList":
								Reusable.Select_Value_DropDown(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								break;
							case "CheckWebElement":
								Reusable.Check_WebElement(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								break;
							case "Checkbox":
								Reusable.Click_CheckBox(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								break;
							case "Radio":
								Reusable.Click_RadioButton(driver,iTestCaseRow, obj, Class_Name,TC_Name);
								break;
							case "ShiftTab":
								 Actions builder = new Actions(driver);
								 Action seriesofaction = builder.keyDown(Keys.SHIFT).sendKeys(Keys.TAB).keyUp(Keys.SHIFT).build();
								 seriesofaction.perform();
								 Thread.sleep(1000);							
								break;
							case "RightArrow":
								 Actions builder1 = new Actions(driver);
								 Action seriesofaction1 = builder1.sendKeys(Keys.RIGHT).build();
								 seriesofaction1.perform();
								 Thread.sleep(1000);
								 break;
							case "Tab":
								 Actions builder2 = new Actions(driver);
								 Action seriesofaction2 = builder2.sendKeys(Keys.TAB).build();
								 seriesofaction2.perform();
								 Thread.sleep(1000);							
								break;
							case "Enter":
								 Actions builder3 = new Actions(driver);
								 Action seriesofaction3 = builder3.sendKeys(Keys.ENTER).build();
								 seriesofaction3.perform();
								 Thread.sleep(1000);							
								break;
							case "Space":
								 Actions builder4 = new Actions(driver);
								 Action seriesofaction4 = builder4.sendKeys(Keys.SPACE).build();
								 seriesofaction4.perform();
								 Thread.sleep(1000);							
								break;
								 
							case "Value":
								Reusable.Enter_value_keystrokes(driver,obj,Class_Name,iTestCaseRow);							
								break;
								
							case "VerifyTable":
								Reusable.VerifyWebTable(driver, iTestCaseRow, obj, Class_Name,TC_Name);
								break;
							
							case "FetchAppValue":
								Reusable.VerifyWebTableScreen(driver, iTestCaseRow, obj, Class_Name,TC_Name);
								break;
								
								
							case "Writetable":
								Reusable.writeExcel("C:\\Users\\889128\\Desktop\\Selenium Luna\\selenium", "Datawrite.xlsx", "Sheet1", Constant.stockArr);
								break;
								
							case "Popup":
								
								Reusable.PopUpaction(driver,obj,Class_Name,iTestCaseRow);							
								break;
								
							case "ParentPopup":
								
								Reusable.ParentPopUpaction(driver,obj,Class_Name,iTestCaseRow);							
																						
								break;
							default:
								break;
								
								
							        		
							}
			        		      				        		  
							
						}
												
					}
				
            	}
        				     	               
				Reusable.WebTableApplication_WriteinExcel(driver,TC_Name);
                Reporter.log("SignIn Action is successfully perfomed");
                
        	}
        	
        	catch(Exception e)
        	{
        		ExcelUtils.setExcelFile(Constant.Path_Excel+Constant.File_RunManager, "Run_Manager");		
       		 
        		int val_run=ExcelUtils.getRowContains_manager(TC_Name, 0, "Run_Manager");
       		 
       		 	ExcelUtils.setCellData_Manager("FAIL", val_run, 2);
        		System.out.println("Error in Login :"+e.toString());
        		 Extend_Report.AddReport("Error in Login Page ", "Fail");
        		 
        		 Extend_Report.AddReport(e.toString(), "Fail");
        		
        	}
        	
        }
    }
