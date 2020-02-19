package com.github.GandhiTC.java.PoiDriver.tests;



import org.apache.commons.lang3.time.DurationFormatUtils;
import com.github.GandhiTC.java.PoiDriver.utilities.PoiDriver;



public class AddColumnTest
{
	public static void main(String[] args)
	{
		long StartTime = System.currentTimeMillis();
		
		
		//	Test 1
		//	Open file to selected sheet
		System.out.println("Test 1");
		System.out.println("Open a file using a file path and switch to a sheet by name (switching sheets by index also available).");
		PoiDriver	poiDriver	= new PoiDriver("src/test/resources/DemoData - Copy.xlsx", "Demo");
		System.out.println("\r\n");
		
		
		//	Test 2
		//	Insert a column
		System.out.println("Test 2");
		System.out.println("Add new column into selected sheet (current sheet if none selected).");
		System.out.println("Numerous overrides planned to be added.");
		poiDriver.addColumn("d");
		System.out.println("\r\n");
		
		
		
		long	EndTime		= System.currentTimeMillis();
		long	Duration	= EndTime - StartTime;
		String	TimeTaken	= DurationFormatUtils.formatDuration(Duration, "HH 'hours' mm 'minutes' ss 'seconds'");
		
		System.out.println("\n\nTotal time to complete: " + TimeTaken);
		System.out.println("\r\nCheck out the .xlsx file to see the changes made, afterwards, close the file and try running RemoveColumnTest.");
	}
}
