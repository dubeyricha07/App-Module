package appModules;
import java.io.FileInputStream;
import java.util.regex.Pattern;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.Reporter;

import pageObjects.BaseClass;
import utility.Constant;
import utility.ExcelUtils;
import utility.Extend_Report;
import utility.Reusable;
     
   
    public class Screen25 {
    	
    	  private static XSSFSheet ExcelWSheet;
          private static XSSFWorkbook ExcelWBook;
          private static XSSFCell Cell;
          private static XSSFRow Row;
          static int noOfColumns;
          static String SheetName;
         
        public  static void Execute(int iTestCaseRow,String TC_Name) throws Exception{
        	WebDriver driver=BaseClass.driver;
        	String class_name1=Thread.currentThread().getStackTrace()[1].getClassName();
        	String[] parts1 = class_name1.split(Pattern.quote("."));
        	//class_name =parts1[0];
        	//String[] parts2 = class_name1.split(Pattern.quote("."));
        	class_name1=parts1[1];
        	String[] parts2 = class_name1.split(Pattern.quote("_"));
        	String class_name=parts2[0];
        	
        	try{      		
        	     
        		FileInputStream ExcelFile = new FileInputStream(Constant.Path_Excel + Constant.File_TestData);
            	//System.out.println(Constant.Path_Excel + Constant.File_TestData);
                // Access the required test data sheet
        		//******CHnages made by richa to implement screen********
        		
        		
                ExcelWBook = new XSSFWorkbook(ExcelFile);
                
             //   List<String> sheetnames = new ArrayList<String>();
                
                for (int i =0; i<ExcelWBook.getNumberOfSheets(); i++) {
                	
                	SheetName = ExcelWBook.getSheetName(i);
                	if(SheetName.contains(class_name)){
                		ExcelWSheet = ExcelWBook.getSheet(SheetName);
                		noOfColumns = ExcelWSheet.getRow(0).getPhysicalNumberOfCells();     
                		break;
                	}
                                	
                }
              
            	                	
            	
				for (int i=1;i < noOfColumns;i++)
            	{
            	
					String  val=ExcelWSheet.getRow(0).getCell(i).getStringCellValue();
					
					int Valpos = val.indexOf("-");
					
					if (Valpos > 0)
					{
						
						String obj=val;
						
						String[] parts = val.split(Pattern.quote("-"));
						
						String Class_Name1 = parts[0];
						String[] parts3 = Class_Name1.split(Pattern.quote("_"));
						String Class_Name = parts3[0];
						String Newword = parts3[1];
						obj = obj.replace("_"+Newword, "");
						String Object_Type = parts[1];
						
						if (Class_Name.equalsIgnoreCase(class_name))
						{
							 
							

							switch (Object_Type){
							
							case "Edit":
								
								Reusable.Enter_Value_WebElement(driver,iTestCaseRow,obj, class_name,TC_Name);
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
								Reusable.ShiftTab_Function(driver,iTestCaseRow, obj, Class_Name,TC_Name); 							
								break;
							case "RightArrow":
								Reusable.RightArrow_Function(driver,iTestCaseRow, obj, Class_Name,TC_Name);	
								 break;
							case "Tab":
								Reusable.Tab_Function(driver,iTestCaseRow, obj, Class_Name,TC_Name);							
								break;
							case "Enter":
								Reusable.Enter_Function(driver,iTestCaseRow, obj, Class_Name,TC_Name);
												
								break;
								
							case "Space":
								Reusable.Space_Function(driver,iTestCaseRow, obj, Class_Name,TC_Name);
												
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
							
							case "iframe":
								Reusable.iframeaction(driver, obj,Class_Name, iTestCaseRow);
								break;
								
							case "Diframe":
								Reusable.Defaultframe(driver, obj,Class_Name, iTestCaseRow);
								break;
								
							case "Writetable":
								Reusable.writeExcel(Constant.currentDir +"//src", "Datawrite.xlsx", "Sheet1", Constant.stockArr);
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
