package com.github.GandhiTC.java.PoiDriver.tests;



import java.io.IOException;
import org.apache.commons.lang3.time.DurationFormatUtils;
import com.github.GandhiTC.java.PoiDriver.utilities.PoiDriver;



public class PredefinedColorsTest
{
	public static void main(String[] args) throws IOException
	{
		long		startTime	= System.currentTimeMillis();

		PoiDriver	poiDriver	= new PoiDriver("src/test/resources/DemoData - Copy.xlsx", "Sample");
		poiDriver.testHSSFColors("F2");
		poiDriver.testXSSFColors("H2");
		poiDriver.selectSheetByName("Sample");

		long		endTime		= System.currentTimeMillis();
		long		duration	= endTime - startTime;
		String		timeTaken	= DurationFormatUtils.formatDuration(duration, "HH 'hours' mm 'minutes' ss 'seconds'");

		System.out.println("\n\nTotal time to complete: " + timeTaken);
		System.out.println("\r\nSee the \"Sample\" sheet for results.");
	}
}