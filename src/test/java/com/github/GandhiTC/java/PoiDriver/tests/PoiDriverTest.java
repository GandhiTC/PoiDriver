package com.github.GandhiTC.java.PoiDriver.tests;



import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.commons.lang3.time.DurationFormatUtils;
import com.github.GandhiTC.java.PoiDriver.utilities.PoiDriver;
import com.github.GandhiTC.java.PoiDriver.utilities.enums.TestCaseResult;



public class PoiDriverTest
{
	public static void main(String[] args)
	{
		long StartTime =  System.currentTimeMillis();
		
		
		//	Test 1
		//	Open file to selected sheet
		System.out.println("Test 1");
		System.out.println("Open a file using a file path and switch to a sheet by name (switching sheets by index also available).");
		PoiDriver	poiDriver	= new PoiDriver("src/test/resources/DemoData - Copy.xlsx", "TestData");
		System.out.println("\r\n");
		
		
		//	Test 2
		//	Get cell value by column and row numbers (not indexes)
		System.out.println("Test 2");
		System.out.println("Find a cell by its column and row numbers (NOT indexes, ie: E8 = (5, 8)), then print its value.");
		String				cellData	= poiDriver.getCellData(5, 8);
		System.out.println("\t - " + cellData);
		System.out.println("\r\n");
		
		
		//	Test 3
		//		using getRowData(String columnHeader, String columnValue)
		//	1)	Select the column whose header is: String columnHeader
		//	2)	From that column, select the cell whose value is: String columnValue
		//	3)  Print the values of every cell that belongs to the same row as that cell.
		System.out.println("Test 3");
		System.out.println("Example scenario:  Print all the values collected by a certain test case.");
		System.out.println("  3a)  Find a column by its column header.  (For example: \"ListOfTestCases\")");
		System.out.println("  3b)  Find a cell in that column by its value.  (For example: \"NameOfSelectedTestCase\")");
		System.out.println("  3c)  Print values of all cells that are on the same row as that cell.");
		ArrayList<String>	data		= poiDriver.getRowData("4", "8");
		String				testcase	= data.get(0).toString();
		String				data1		= data.get(1).toString();
		String				data2		= data.get(2).toString();
		String				baseURL		= data.get(5).toString();
		String				method		= data.get(6).toString();
		System.out.println("\t - " + testcase);
		System.out.println("\t - " + data1);
		System.out.println("\t - " + data2);
		System.out.println("\t - " + data.get(3));
		System.out.println("\t - " + data.get(4));
		System.out.println("\t - " + baseURL);
		System.out.println("\t - " + method);
		System.out.println("\r\n");
		
		//	Examples of how data could be used with selenium
//		driver.get("http://" + baseURL);
//		driver.findElement(By.xpath("//input[@id='data1_input']").sendKeys(data1);
//		driver.findElement(By.xpath("//input[@id='TestCases_input']").sendKeys(data.get(3));
		
		
		//	Test 4
		//	Mark test cases as pass/fail
		System.out.println("Test 4");
		System.out.println("Mark test cases as pass/fail, check file afterwards.");
			//	4a - Individually & immediately mark test results as pass/fail by cell address/reference
			System.out.println("  4a)  Individually and immediately - by cell address/reference.");
			poiDriver.markTestCaseResult("J4", TestCaseResult.FAIL);
			poiDriver.markTestCaseResult("J5", TestCaseResult.PASS);
			poiDriver.markTestCaseResult("J10", TestCaseResult.PASS);
			
			//	4b - Use a ConcurrentHashMap to hold test results by cell address/reference, then mark them all later
			System.out.println("  4b)  Store test results, then mark them all later - by cell address/reference.");
			ConcurrentHashMap<String, TestCaseResult> testResults = new ConcurrentHashMap<String, TestCaseResult>();
				testResults.put("J6", TestCaseResult.FAIL);
				testResults.put("J7", TestCaseResult.SKIPPED);
				testResults.put("J8", TestCaseResult.FAIL);
			poiDriver.markTestCaseResults(testResults);
			
			//	4c - Define column headers ahead of time, then use name of test case to mark as pass/fail
			System.out.println("  4c)  Individually and immediately - by setting column headers ahead of time, then later using name of test case.");
			poiDriver.setTestColumnsHeaders("TestCase", "Result", true);
			poiDriver.markTestCaseResultByName("Purchase", TestCaseResult.PASS);
			poiDriver.markTestCaseResultByName("Add Profile", TestCaseResult.FAIL);
			
			//	4d - Use a ConcurrentHashMap to hold test results by test case name, then mark them all later
			System.out.println("  4d)  Store test results, then mark them all later - by setting column headers ahead of time, then later using name of test case.");
			ConcurrentHashMap<String, TestCaseResult> testResultsByName = new ConcurrentHashMap<String, TestCaseResult>();
				testResultsByName.put("Login", TestCaseResult.FAIL);
				testResultsByName.put("Delete Profile", TestCaseResult.SKIPPED);
				testResultsByName.put("Get Count", TestCaseResult.FAIL);
				testResultsByName.put("Final Step", TestCaseResult.PASS);
			poiDriver.markTestCaseResultsByName(testResultsByName);
		System.out.println("\r\n");
		
		
		//	Test 5
		//	Switch to another sheet by sheet index
		System.out.println("Test 5");
		System.out.println("Switch to a different sheet by index. (Switching by sheet name is also available and preferred.)");
		poiDriver.selectSheetAtIndex(2);
		System.out.println("\r\n");
		
		
		//	Test 6
		//	Set cell value by cell reference
		System.out.println("Test 6");
		System.out.println("Select cells by their cell address/reference (examples: \"B2\", \"D4\") and set their values.");
			//	string
			poiDriver.setCellData("B2", "This is \nB2 \nis This");
			String lines[] = poiDriver.getCellData("B2").split("\\r?\\n");
			boolean firstLine = true;
			for(String line : lines)
			{
				if(firstLine)
				{
					System.out.println("\t - " + line);
					firstLine = false;
				}
				else
				{
					System.out.println("\t   " + line);
				}
			}
			//	date
			poiDriver.setCellData("d4", new Date().toString());		//	Calendar.getInstance().getTime().toString()
			System.out.println("\t - " + poiDriver.getCellData("d4"));
		System.out.println("\r\n");
			
			
		//	Test 7
		//	Set cell value by row number, column number
		System.out.println("Tests 7");
		System.out.println("Select cells by their row and column numbers (not indexes), and set their values in specific formats.");
		System.out.println("\tCheck file afterwards to confirm cell formatting.");
			//	double
			poiDriver.setCellData(6, 6, "1.1234567890123456");
			System.out.println("\t - " + poiDriver.getCellData(6, 6));
			//	long
			poiDriver.setCellData(8, 8, "7000000000");
			System.out.println("\t - " + poiDriver.getCellData(8, 8));
			//	percentage
			poiDriver.setCellData(10, 10, "25.7535%");
			System.out.println("\t - " + poiDriver.getCellData(10, 10));
			//	float
			poiDriver.setCellData(12, 12, "696.969");
			System.out.println("\t - " + poiDriver.getCellData(12, 12));
			//	currency
			poiDriver.setCellData(14, 14, "$12.34");
			System.out.println("\t - " + poiDriver.getCellData(14, 14));
			//	int
			poiDriver.setCellData(16, 16, "2500");
			System.out.println("\t - " + poiDriver.getCellData(16, 16));
		System.out.println("\r\n");
		
		
		//	Test 8
		//	Set custom cell style
		//	Completed during Test 7
		System.out.println("Test 8");
		System.out.println("Set custom cell style, check file to confirm.");
		System.out.println("\r\n");
		
		
		//	Test 9
		//	create a new sheet
		System.out.println("Test 9");
		System.out.println("Add new sheets to the workbook.");
		System.out.println("  9a)  Create a new sheet called \"IndexedSheet\" and place it in the 2nd position.");
		poiDriver.addSheetAtIndex("IndexedSheet", 1);
		System.out.println("\t - New sheet \"IndexedSheet\" created, check file to confirm.");
		System.out.println("  9b)  Create a new sheet called \"LastSheet\", it will automatically be placed in the last position.");
		poiDriver.addSheet("LastSheet");
		System.out.println("\t - New sheet \"LastSheet\" created, check file to confirm.");
		System.out.println("\r\n");
		
		
		
		long	EndTime		= System.currentTimeMillis();
		long	Duration	= EndTime - StartTime;
		String	TimeTaken	= DurationFormatUtils.formatDuration(Duration, "HH 'hours' mm 'minutes' ss 'seconds'");
		
		System.out.println("\n\nTotal time to complete: " + TimeTaken);
		System.out.println("\r\nCheck out the .xlsx file to see the changes made, afterwards, close the file and try running AddColumnTest.");
		System.err.println("\r\nNOTE:  Selecting sheets, columns, and/or rows by their index value is not recommended.");
		System.err.println("\tAfter viewing the .xlsx file, close it, and run this test again to see why selecting by index is not recommended.");
	}
}
