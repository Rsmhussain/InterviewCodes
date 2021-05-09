//Java
import java.io.*;
import java.util.*;

//Selenium Jars
import org.openqa.selenium.*;
import org.openqa.selenium.interactions.*;
import org.openqa.selenium.chrome.ChromeDriver;

//Apache
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelAutomation
{
	public static void main(String args[]) throws Exception
	{
		try 
		{
			
		
		System.out.println("Welcome");
		File src= new File("G:\\Aslam Personal\\Interview Preparation\\Interview Programs\\AutomationData.Xlsx");
		FileInputStream fis=new FileInputStream(src);
		XSSFWorkbook wb= new XSSFWorkbook(fis);
		XSSFSheet sheet= wb.getSheetAt(0);
		System.out.println("The Sheet Name:"+sheet);
		int rowcount=sheet.getLastRowNum()+1;
		System.out.println("Toal Row Cocunt:"+rowcount);
		XSSFRow row=sheet.getRow(0);
		int cellcount=row.getLastCellNum();
		System.out.println("Toal Column Cocunt:"+cellcount);
		
		for(int i=1;i<rowcount;i++)
		{
			for(int j=1;j<cellcount;j++)
			{
			if(j<2)
			{
			String username=sheet.getRow(i).getCell(j).getStringCellValue();
			System.out.println("User Name: "+username);
			}
			else if(j>=2)
			{
			int password=(int) sheet.getRow(i).getCell(j).getNumericCellValue();
			System.out.println("Password : "+password);
			}
			}
		}
		}
		
		//Last Update 
		    
		catch(Exception e)
		{
			System.out.println(e);
		}
				
	
	}

}
