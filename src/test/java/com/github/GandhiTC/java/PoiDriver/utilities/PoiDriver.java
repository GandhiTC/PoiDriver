package com.github.GandhiTC.java.PoiDriver.utilities;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.AbstractMap.SimpleEntry;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.validator.routines.BigDecimalValidator;
import org.apache.commons.validator.routines.CurrencyValidator;
import org.apache.commons.validator.routines.DateValidator;
import org.apache.commons.validator.routines.DoubleValidator;
import org.apache.commons.validator.routines.PercentValidator;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.SkipException;
import com.github.GandhiTC.java.PoiDriver.utilities.enums.ErrorHandling;
import com.github.GandhiTC.java.PoiDriver.utilities.enums.PrintToConsole;
import com.github.GandhiTC.java.PoiDriver.utilities.enums.RunMode;
import com.github.GandhiTC.java.PoiDriver.utilities.enums.TestCaseResult;




/*
	//  List of To-Do's
	//	TODO:	add addRow methods
	//	TODO:	add removeRow methods
	//	TODO:	add getCellData methods for selected sheet by name & by index
	//	TODO:	add setCellData methods for selected sheet by name & by index
	//	TODO:	finish implementing checking run mode
	//	TODO:	implement a TestNG listener or extent reports listener to update pass/fail values in excel automatically
	//	TODO:	look into adding another layer of overloading for all methods, to allow selecting cell by row number and column letter separately
	//	TODO:	add numerous overrides to all methods, including addColumn, that allows setting value of a cell in that column, not just a header
	//	TODO:	look into breaking up code and grouping into separate classes/interfaces
	//	TODO:	look into implementing builder patterns
*/
 



public class PoiDriver
{
	//	Class vars
	public  ErrorHandling	onError				= ErrorHandling.ContinueOnError;
	public	PrintToConsole	consoleErrorType	= PrintToConsole.StackTrace;
	
	private FileInputStream	fis;
	private XSSFWorkbook	workbook;
	private XSSFSheet		currentSheet;		//	it is like ActiveSheet in POI, but specifically for PoiDriver
	private DataFormatter	formatter;
	private String			filePath;
	private XSSFColor		limeGreen 			= new XSSFColor(new byte[] { (byte)50, (byte)205, (byte)50 }, null);
	
	private ConcurrentHashMap<SimpleEntry<String, String>, SimpleEntry<String, String>> columnCache = new ConcurrentHashMap<SimpleEntry<String, String>, SimpleEntry<String, String>>();
	
	
	
	
	//	Constructors
	public PoiDriver(String filePath)
	{
		openFile(filePath);
	}
	
	
	public PoiDriver(String filePath, String sheetName)
	{
		openFile(filePath, sheetName);
	}
	
	
	public PoiDriver(String filePath, int sheetIndex)
	{
		openFile(filePath, sheetIndex);
	}
	
	
	
	
	//	Exception handling helper
	private void onErrorDo(Exception e)
	{
		if(consoleErrorType == PrintToConsole.Message)
		{
			System.err.println(e.getMessage());
		}
		else
			if(consoleErrorType == PrintToConsole.StackTrace)
			{
				e.printStackTrace();
			}
		
		
		if(onError == ErrorHandling.ContinueOnError)
		{
			//	nothing to do here
		}
		else
			if(onError == ErrorHandling.StopTestingOnError)
			{
				//	TODO:	Add code to send TestNG SkipException() to remaining tests
			}
		else
			if(onError == ErrorHandling.TerminateOnError)
			{
				//	TODO:	Add code to send System.exit(1) to non-test scripts
			}
		
		
		//	TODO:	Add code to send full stack trace message to logs
	}
	
	
	
	
	//	Open a workbook file and set values for class variables
	private void openFile(String filePath)
	{
		try
		{
			File file = new File(filePath.trim());
			
			if(validateFile(file) == true)
			{
				this.filePath	= filePath;
				fis				= new FileInputStream(file);
				workbook		= new XSSFWorkbook(fis);
				formatter		= new DataFormatter();
				
				if(workbook.getNumberOfSheets() < 1)
				{
					System.err.println("Sheet not selected./r/nCurrent workbook does not have any sheets!\r\nPlease add data to the workbook before continuing.");
					
					this.filePath	= null;
					fis				= null;
					workbook		= null;
					formatter		= null;
					currentSheet	= null;
					
					return;
				}
				else
				{
					int index = workbook.getActiveSheetIndex();
					
					workbook.setActiveSheet(index);
					workbook.setSelectedTab(index);
					currentSheet = workbook.getSheetAt(index);
					
					currentSheet.lockFormatColumns(false);
				}
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
	}
	
	
	private void openFile(String filePath, String sheetName)
	{
		try
		{
			File file = new File(filePath.trim());
			
			if(validateFile(file) == true)
			{
				this.filePath	= filePath;
				fis				= new FileInputStream(file);
				workbook		= new XSSFWorkbook(fis);
				formatter		= new DataFormatter();
				
				int index = workbook.getSheetIndex(sheetName.trim());
				
				if(index == -1)
				{
					System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet by this name.");
					
					this.filePath	= null;
					fis				= null;
					workbook		= null;
					formatter		= null;
					currentSheet	= null;
					
					return;
				}
				else
				{
					workbook.setActiveSheet(index);
					workbook.setSelectedTab(index);
					currentSheet = workbook.getSheetAt(index);
					
					currentSheet.lockFormatColumns(false);
				}
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
	}
	
	
	private void openFile(String filePath, int sheetIndex)
	{
		try
		{
			File file = new File(filePath.trim());
			
			if(validateFile(file) == true)
			{
				this.filePath	= filePath;
				fis				= new FileInputStream(file);
				workbook		= new XSSFWorkbook(fis);
				formatter		= new DataFormatter();
				
				if(workbook.getSheetAt(sheetIndex) == null)
				{
					System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
					
					this.filePath	= null;
					fis				= null;
					workbook		= null;
					formatter		= null;
					currentSheet	= null;
					
					return;
				}
				else
				{
					workbook.setActiveSheet(sheetIndex);
					workbook.setSelectedTab(sheetIndex);
					currentSheet = workbook.getSheetAt(sheetIndex);
					
					currentSheet.lockFormatColumns(false);
				}
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
	}
	
	
	
	
	//	Validate workbook file
	private boolean validateFile(File file)
	{
		if(!file.isFile())
		{
			if(!file.exists())
			{
				String	fileName	= file.getName();
				int		dotIndex	= fileName.lastIndexOf('.');
				
				if(dotIndex == -1)
				{
					System.err.println("Please make sure you have selected a file, not a path.");
					return false;
				}
				
				String	extension	= fileName.substring(dotIndex + 1);
				
				if(!"xlsx".equalsIgnoreCase(extension))
				{
					System.err.println("Please make sure the file extension is: .xlsx");
					return false;
				}
				
				System.err.println("File does not exist.");
				return false;
			}
			else
			{
				System.err.println("Error in given file path.");
				return false;
			}
		}
		else
		{
			return true;
		}
	}
	
	
	
	
	//	Close file
	private void closeFile()
	{
		try
		{
			if(workbook != null)
			{
				workbook.close();
			}
			
			if(fis != null)
			{
				fis.close();
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
	}
	
	
	
	
	//	Add a new sheet
	public void addSheet(String sheetName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			String safeName = WorkbookUtil.createSafeSheetName(sheetName.trim());
			workbook.createSheet(safeName);
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void addSheetAtIndex(String sheetName, int index)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			String safeName = WorkbookUtil.createSafeSheetName(sheetName.trim());
			workbook.createSheet(safeName);
			workbook.setSheetOrder(safeName, index);
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Check if sheet exists
	public boolean sheetExists(String sheetName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			for(int x = 0; x <= workbook.getNumberOfSheets(); x++)
			{
				if(sheetName.trim().equalsIgnoreCase(workbook.getSheetName(x)))
				{
					return true;
				}
			}
			
			return false;
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return false;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Remove a sheet
	public void removeSheetByName(String sheetName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			int index = workbook.getSheetIndex(sheetName.trim());
			
			if(index == -1)
			{
				System.err.println("Unable to remove selected sheet.\r\nCurrent workbook does not have a sheet by this name.");
			}
			else
			{
				workbook.removeSheetAt(index);
				
				OutputStream fileOut = new FileOutputStream(filePath);
				workbook.write(fileOut);
				fileOut.close();
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void removeSheetAtIndex(int sheetIndex)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			if(workbook.getSheetAt(sheetIndex) == null)
			{
				System.err.println("Sheet not removed./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			else
			{
				workbook.removeSheetAt(sheetIndex);
				
				OutputStream fileOut = new FileOutputStream(filePath);
				workbook.write(fileOut);
				fileOut.close();
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Select a sheet
	public void selectSheetByName(String sheetName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			int index = workbook.getSheetIndex(sheetName.trim());
			
			if(index == -1)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet by this name.");
				return;
			}
			else
			{
				workbook.setActiveSheet(index);
				workbook.setSelectedTab(index);
				currentSheet = workbook.getSheetAt(index);
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void selectSheetAtIndex(int sheetIndex)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			if(workbook.getSheetAt(sheetIndex) != null)
			{
				workbook.setActiveSheet(sheetIndex);
				workbook.setSelectedTab(sheetIndex);
				currentSheet = workbook.getSheetAt(sheetIndex);
			}
			else
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Display current sheet name
	public String getCurrentSheet()
	{
		return currentSheet.getSheetName();
	}
	
	
	
	
	//	Get row count in selected sheet, default false
	public int getRowCount(boolean getPhysicalCount)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			return getPhysicalCount == true ? currentSheet.getPhysicalNumberOfRows() : currentSheet.getLastRowNum() + 1;
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public int getRowCount(String sheetName, boolean getPhysicalCount)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			int index = workbook.getSheetIndex(sheetName);

			if(index == -1)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet by this name.");
				return 0;
			}
			else
			{
				XSSFSheet	sheet  = workbook.getSheetAt(index);
				int 		number = getPhysicalCount == true ? sheet.getPhysicalNumberOfRows() : sheet.getLastRowNum() + 1;
				
				return number;
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public int getRowCount(int sheetIndex, boolean getPhysicalCount)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			
			if(sheet != null)
			{
				return getPhysicalCount == true ? sheet.getPhysicalNumberOfRows() : sheet.getLastRowNum() + 1;
			}
			else
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return 0;
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	returns number of columns in a sheet
	public int getMaxColumnCount()
	{
		StackTraceElement[]	stacktrace		= Thread.currentThread().getStackTrace();
		StackTraceElement	element			= stacktrace[2];
//		String				methodName		= element.getMethodName();
//		boolean				isClassCalled	= methodName.equalsIgnoreCase("processAddingColumn") && element.getClassName().equalsIgnoreCase(this.getClass().getName());
		boolean				isClassCalled	= element.getClassName().equalsIgnoreCase(this.getClass().getName());
		
		try
		{
			if(!isClassCalled)
			{
				openFile(filePath, currentSheet.getSheetName());
			}
			
			int	maxColumns	= 0;
			int	firstRowNum	= currentSheet.getFirstRowNum();
			int	lastRowNum	= currentSheet.getLastRowNum();
			
			for(int r = firstRowNum; r <= lastRowNum; r++)
			{
				Row row = currentSheet.getRow(r);
				
				if(row == null)
				{
					continue;
				}
				
				maxColumns = Math.max(maxColumns, row.getLastCellNum());
			}
			
//			System.out.println("Called by method: " + methodName);
//			System.out.println("Called by class : " + element.getClassName());
//			System.out.println("This class name : " + this.getClass().getName());
//			System.out.println("isClassCalled   = " + isClassCalled);
//			System.out.println("maxColumns      = " + maxColumns);
			
			return Math.max(0, maxColumns - 1);
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			if(!isClassCalled)
			{
				closeFile();
			}
		}
	}
	
	
	public int getHeaderCount()
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			Row firstRow = currentSheet.getRow(currentSheet.getFirstRowNum());
			
			if(firstRow == null)
			{
				return 0;
			}
			
			if(firstRow.getLastCellNum() == -1)
			{
				return 0;
			}
			else
			{
				return Math.max(0, firstRow.getLastCellNum() - 1);
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public int getHeaderCount(String sheetName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			XSSFSheet sheet = workbook.getSheet(sheetName);
			
			if(sheet == null)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet by this name.");
				return 0;
			}
			
			Row firstRow = sheet.getRow(sheet.getFirstRowNum());
			
			if(firstRow == null)
			{
				return 0;
			}
			
			if(firstRow.getLastCellNum() == -1)
			{
				return 0;
			}
			else
			{
				return Math.max(0, firstRow.getLastCellNum() - 1);
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public int getHeaderCount(int sheetIndex)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			
			if(sheet == null)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return 0;
			}
			else
			{
				Row firstRow = sheet.getRow(sheet.getFirstRowNum());
				
				if(firstRow == null)
				{
					return 0;
				}
				
				if(firstRow.getLastCellNum() == -1)
				{
					return 0;
				}
				else
				{
					return Math.max(0, firstRow.getLastCellNum() - 1);
				}
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return 0;
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Insert a new column
	public void addColumn(String columnLetter)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			int column_index  = CellReference.convertColStringToIndex(columnLetter.toUpperCase());
			
			if(processAddingColumn(currentSheet, column_index, "", true) == false)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			
			autoResize(currentSheet.getSheetName(), column_index);
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void addColumn(String columnLetter, String columnHeaderName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			int column_index  = CellReference.convertColStringToIndex(columnLetter);
			
			if(processAddingColumn(currentSheet, column_index, columnHeaderName, false) == false)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void addColumn(int columnNumber, String columnHeaderName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			if(processAddingColumn(currentSheet, columnNumber - 1, columnHeaderName, false) == false)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void addColumn(String sheetName, int columnNumber, String columnHeaderName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			if(processAddingColumn(workbook.getSheet(sheetName), columnNumber - 1, columnHeaderName, false) == false)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void addColumn(int sheetIndex, int columnNumber, String columnHeaderName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			if(processAddingColumn(workbook.getSheetAt(sheetIndex), columnNumber - 1, columnHeaderName, false) == false)
			{
				System.err.println("Sheet not selected./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	private boolean processAddingColumn(XSSFSheet sheet, int columnIndex, String columnHeaderName, boolean usingMaxColumns)
	{
		if(sheet == null)
		{
			return false;
		}
		else
		{
			int		cellCount	= 0;
			int		firstRowNum	= sheet.getFirstRowNum();
			Row		firstRow	= sheet.getRow(firstRowNum);
			Cell	cell		= null;
			
			if(firstRow == null)
			{
				firstRow 	= sheet.createRow(0);
			}

			if(firstRow.getLastCellNum() == -1)
			{
				cell 		= firstRow.createCell(0);
			}
			else
			{
				if(columnIndex < 0)
				{
					cell 		= firstRow.createCell(firstRow.getLastCellNum());
				}
				else
				{
					cellCount 	= usingMaxColumns == true ? Math.max(getMaxColumnCount(), 1) : Math.max(getHeaderCount(), 1);
					
					sheet.shiftColumns(columnIndex, cellCount, 1);
					
					cell 		= firstRow.createCell(columnIndex);
				}
			}
			
			if(columnHeaderName != "")
			{
				cell.setCellValue(columnHeaderName);
			}
			
			if(cell.getColumnIndex() > 0)
			{
				Cell tmpCell = firstRow.getCell(cell.getColumnIndex() - 1);
				
				if(tmpCell != null)
				{

					if(tmpCell.getCellStyle() != null)
					{
						cell.setCellStyle(tmpCell.getCellStyle());
					}
				}
			}
			
			for(int c = Math.max(1, columnIndex); c <= cellCount; c++)
			{
				sheet.setColumnWidth(c, sheet.getColumnWidth(c - 1));
			}
			
			return true;
		}
	}
	
	
	
	
	//	Remove a column
	public void removeColumn(String columnLetter)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			int column_index  = CellReference.convertColStringToIndex(columnLetter.toUpperCase());
			
			processRemovingColumn(currentSheet, column_index);
			
			autoResize(currentSheet.getSheetName(), column_index);
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void removeColumn(int columnIndex)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			processRemovingColumn(currentSheet, columnIndex);
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void removeColumn(String sheetName, int columnIndex)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			XSSFSheet sheet = workbook.getSheet(sheetName);
			
			if(sheet == null)
			{
				System.err.println("Column not removed./r/nCurrent workbook does not have a sheet by this name.");
				return;
			}
			
			processRemovingColumn(currentSheet, columnIndex);
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void removeColumn(int sheetIndex, int columnIndex)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			
			if(sheet == null)
			{
				System.err.println("Column not removed./r/nCurrent workbook does not have a sheet at this index.");
				return;
			}
			
			processRemovingColumn(currentSheet, columnIndex);
			
			OutputStream fileOut = new FileOutputStream(filePath);
			workbook.write(fileOut);
			fileOut.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	private void processRemovingColumn(XSSFSheet sheet, int columnIndex)
	{
		int maxColumn = 0;
		
		for(int r = 0; r <= sheet.getLastRowNum(); r++)
		{
			Row row = sheet.getRow(r);
			
			if(row == null)
			{
				continue;
			}
			
			int lastColumn = row.getLastCellNum();
			
			if(lastColumn > maxColumn)
			{
				maxColumn = lastColumn;
			}
			
			if(lastColumn < columnIndex)
			{
				continue;
			}

			for(int x = columnIndex + 1; x <= lastColumn; x++)
			{
				Cell oldCell = row.getCell(x - 1);
				
				if(oldCell != null)
				{
					row.removeCell(oldCell);
				}
				
				Cell nextCell = row.getCell(x);

				if(nextCell != null)
				{
					Cell newCell = row.createCell(x - 1, nextCell.getCellType());
					cloneCell(newCell, nextCell);
				}
			}
		}

		// Adjust the column widths
		for(int c = 0; c < maxColumn; c++)
		{
			sheet.setColumnWidth(c, sheet.getColumnWidth(c + 1));
		}
	}
	
	
	private void cloneCell(Cell newCell, Cell oldCell)
	{
		newCell.setCellComment(oldCell.getCellComment());
		newCell.setCellStyle(oldCell.getCellStyle());
		
		switch(newCell.getCellType())
		{
			case BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;

			case NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;

			case STRING:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;

			case ERROR:
				newCell.setCellValue(oldCell.getErrorCellValue());
				break;

			case FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
				
			case BLANK:
				newCell.setBlank();
				break;
				
			default:
				newCell.setBlank();
				break;
		}
	}
	
	
	
	
	//	Column assignments - Used in markTestCaseResultByName() & markTestCaseResultsByName(), but could also be used in getRowData() with a little editing
	public void setTestColumnsHeaders(String testCaseColumnHeader, String testResultColumnHeader, boolean forCurrentSheetOnly)
	{
		SimpleEntry<String, String> tmpKey = new SimpleEntry<String, String>(new File(filePath).getName(), currentSheet.getSheetName());
		SimpleEntry<String, String> tmpVal = new SimpleEntry<String, String>(testCaseColumnHeader, testResultColumnHeader);
		
		if(columnCache.containsKey(tmpKey))
		{
			columnCache.replace(tmpKey, tmpVal);
		}
		else
		{
			columnCache.put(tmpKey, tmpVal);
		}
			
		if(forCurrentSheetOnly == false)
		{
			Iterator<Entry<SimpleEntry<String, String>, SimpleEntry<String, String>>> entries = columnCache.entrySet().iterator();
			
			while(entries.hasNext())
			{
				Entry<SimpleEntry<String, String>, SimpleEntry<String, String>>	entry = entries.next();
				entry.setValue(tmpVal);
			}
		}
		
		tmpKey = null;
		tmpVal = null;
	}
	
	
	
	
	//	ArrayList which returns row data of selected test case
	public ArrayList<String> getRowData(String columnHeader, String columnValue)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			ArrayList<String>	arrayList	= new ArrayList<String>();
			Iterator<Row>		rows		= currentSheet.iterator();	//	A spreadsheet is a collection of rows
			Row					firstRow	= rows.next();				//	The first row to have data, **assumes this is the row with headers**
			Iterator<Cell>		cells		= firstRow.cellIterator();	//	A row is a collection of cells
			int					x			= firstRow.getFirstCellNum();
			int					columnNum	= x;
	
			//	Identify the column number whose header is the columnHeader parameter
			while(cells.hasNext())
			{
				Cell	currentCell	= cells.next();
				String	cellData	= formatter.formatCellValue(currentCell);
	
				if(cellData.equalsIgnoreCase(columnHeader))
				{
					columnNum = x;
					break;
				}
	
				x++;
			}
	
			//	Scan the column to identify the cell whose value is the columnValue parameter, then make a note of its row
			while(rows.hasNext())
			{
				Row		currentRow	= rows.next();
				Cell	columnCell	= currentRow.getCell(columnNum);
				String	cellData	= formatter.formatCellValue(columnCell);
	
				//	Once row is found, feed data into arrayList
				if(cellData.equalsIgnoreCase(columnValue))
				{
					Iterator<Cell> currentRowsCells = currentRow.cellIterator();
	
					while(currentRowsCells.hasNext())
					{
						Cell currentCell = currentRowsCells.next();
						arrayList.add(formatter.formatCellValue(currentCell));
					}
				}
			}
	
			return arrayList;
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
		
		return new ArrayList<String>();
	}
	
	
	
	
	//	Get cell value
	public String getCellData(int columnNumber, int rowNumber)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			if(rowNumber < 1)
			{
				System.err.println("Unable to get cell data.\r\nPlease enter a valid row number.");
				return "";
			}
			
			if(columnNumber < 1)
			{
				System.err.println("Unable to get cell data.\r\nPlease enter a valid column number.");
				return "";
			}
			
			XSSFRow		row			= currentSheet.getRow(rowNumber - 1);
			
			if(row == null)
			{
				System.err.println("Unable to get cell data.\r\nPlease select a valid row.");
				return "";
			}
			
			XSSFCell	cell		= row.getCell(columnNumber - 1);
			
			if(cell == null)
			{
				System.err.println("Unable to get cell data.\r\nPlease select a valid cell.");
				return "";
			}
			
			String		cellData	= formatter.formatCellValue(cell);
			
			return cellData;
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return "";
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public String getCellData(String cellReference)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			CellReference	ref			= new CellReference(cellReference);
			
			XSSFRow			row			= currentSheet.getRow(ref.getRow());
			
			if(row == null)
			{
				System.err.println("Unable to get cell data.\r\nPlease select a valid row.");
				return "";
			}
			
			XSSFCell		cell		= row.getCell(ref.getCol());
			
			if(cell == null)
			{
				System.err.println("Unable to get cell data.\r\nPlease select a valid cell.");
				return "";
			}
			
			String			cellData	= formatter.formatCellValue(cell);
			
			return cellData;
		}
		catch(Exception e)
		{
			onErrorDo(e);
			return "";
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Set cell value
	public void setCellData(int rowNumber, int columnNumber, String data)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			rowNumber		= rowNumber    - 1;
			columnNumber	= columnNumber - 1;
			
			XSSFRow  row;
			XSSFCell cell;
			
			if(currentSheet.getRow(rowNumber) == null)
			{
				row  = currentSheet.createRow(rowNumber);
			}
			else
			{
				row  = currentSheet.getRow(rowNumber);
			}
			
			if(row.getCell(columnNumber) == null)
			{
				cell = row.createCell(columnNumber);
			}
			else
			{
				cell = row.getCell(columnNumber);
			}
			
			processSettingCellData(rowNumber, columnNumber, data, row, cell);
			
			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void setCellData(String cellReference, String data)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			CellReference	ref				= new CellReference(cellReference);
			int				rowNumber		= ref.getRow();
			int				columnNumber	= ref.getCol();
			
			XSSFRow  row;
			XSSFCell cell;
			
			if(currentSheet.getRow(rowNumber) == null)
			{
				row  = currentSheet.createRow(rowNumber);
			}
			else
			{
				row  = currentSheet.getRow(rowNumber);
			}
			
			if(row.getCell(columnNumber) == null)
			{
				cell = row.createCell(columnNumber);
			}
			else
			{
				cell = row.getCell(columnNumber);
			}
			
			processSettingCellData(rowNumber, columnNumber, data, row, cell);

			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	private void processSettingCellData(int rowNumber, int columnNumber, String data, XSSFRow row, XSSFCell cell)
	{
		XSSFCellStyle	cellStyle	= workbook.createCellStyle();
		CellStyleHelper	csHelper	= new CellStyleHelper(workbook, cellStyle);
		
		csHelper.setFont("Arial", (short)20, IndexedColors.BLUE.getIndex(), true, FontUnderline.DOUBLE_ACCOUNTING, true);
		csHelper.setAllBorderStyles(BorderStyle.MEDIUM_DASH_DOT_DOT, IndexedColors.BROWN.getIndex(), true, true, true, true);
		csHelper.setFill(IndexedColors.GREY_25_PERCENT.getIndex(), FillPatternType.THIN_FORWARD_DIAG);
		
		cell.setCellType(CellType.NUMERIC);
		
		Locale testLocale = Locale.getDefault();
		
		TopLoopLevel:
		if(NumberUtils.isParsable(data))
		{
			DoubleValidator	dblValidator	= DoubleValidator.getInstance();
			Double			dataAsDouble	= dblValidator.validate(data, testLocale);

			if(dataAsDouble != null)
			{
				if((dataAsDouble == Math.floor(dataAsDouble)) && !Double.isInfinite(dataAsDouble))
				{
					cell.setCellValue(new BigInteger(data).longValue());
				}
				else
				{
					cell.setCellValue(dataAsDouble);
					cellStyle.setDataFormat(workbook.createDataFormat().getFormat("#0.00############")); // max decimal places in excel
				}
				
				break TopLoopLevel;
			}
		}
		else
		{
			BigDecimalValidator	bdValidator	= CurrencyValidator.getInstance();
			BigDecimal			amount		= bdValidator.validate(data, testLocale);

			if(amount != null)
			{
				cell.setCellValue(amount.doubleValue());
				cellStyle.setDataFormat(workbook.createDataFormat().getFormat("_($* #,##0.00_);[Red]_($* (#,##0.00);_($* \"-\"??_);_(@_)"));
				break TopLoopLevel;
			}
			else
			{
				String 				pcntFormat	= "#0.00############%";
				
				BigDecimalValidator	pValidator	= PercentValidator.getInstance();
				BigDecimal			percentage	= pValidator.validate(data, pcntFormat, testLocale);

				if(percentage != null)
				{
					cell.setCellValue(percentage.doubleValue());
					cellStyle.setDataFormat(workbook.createDataFormat().getFormat(pcntFormat)); // max decimal places in excel
					break TopLoopLevel;
				}
				else
				{
					DateValidator	dtValidator	= DateValidator.getInstance();
					Date			dateVal		= dtValidator.validate(data, "E MMM dd HH:mm:ss zzz yyyy");
					
					if(dateVal != null)
					{
						cell.setCellValue(dateVal);
						cellStyle.setDataFormat(workbook.createDataFormat().getFormat("ddd mmm dd, yyyy hh:mm:ss AM/PM"));
						break TopLoopLevel;
					}
					else
					{
						cell.setCellType(CellType.STRING);
						cell.setCellValue(data);
						break TopLoopLevel;
					}
				}
			}
		}

		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);
		
		currentSheet.autoSizeColumn(columnNumber);
		row.setHeight((short)-1);
	}
	
	
	
	
	//	Auto resize all rows and columns
	private void autoResize(String sheetName, int columnIndex)
	{
		StackTraceElement[]	stacktrace		= Thread.currentThread().getStackTrace();
		StackTraceElement	element			= stacktrace[2];
//		String				methodName		= element.getMethodName();
//		boolean				isClassCalled	= methodName.equalsIgnoreCase("processAddingColumn") && element.getClassName().equalsIgnoreCase(this.getClass().getName());
		boolean				isClassCalled	= element.getClassName().equalsIgnoreCase(this.getClass().getName());
		
		try
		{
			if(!isClassCalled)
			{
				openFile(filePath, currentSheet.getSheetName());
			}
			
			XSSFSheet	sheet		= workbook.getSheet(sheetName);
			
			if(sheet == null)
			{
				return;
			}
			
			if(sheet.getFirstRowNum() == -1)
			{
				return;
			}
			
			int			cellCount	= getMaxColumnCount();
			
			for(int r = 0; r <= sheet.getLastRowNum(); r++)
			{
				Row rRow = sheet.getRow(r) == null ? sheet.createRow(r) : sheet.getRow(r);

				for(int c = 0; c <= cellCount; c++)
				{
					@SuppressWarnings("unused")
					Cell cCell = rRow.getCell(c) == null ? rRow.createCell(c) : rRow.getCell(c);

					sheet.autoSizeColumn(c);
				}

				rRow.setHeight((short)-1);
			}
			
			if(!isClassCalled)
			{
				OutputStream fileOut = new FileOutputStream(filePath);
				workbook.write(fileOut);
				fileOut.close();
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			if(!isClassCalled)
			{
				closeFile();
			}
		}
	}
	
	
	
	
	//	Place Pass/Fail marker on cell
	public void markTestCaseResult(String cellReference, TestCaseResult testCaseResult)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			CellReference	ref				= new CellReference(cellReference);
			int				rowNumber		= ref.getRow();
			int				columnNumber	= ref.getCol();
			
			XSSFRow  row;
			XSSFCell cell;
			
			if(currentSheet.getRow(rowNumber) == null)
			{
				row  = currentSheet.createRow(rowNumber);
			}
			else
			{
				row  = currentSheet.getRow(rowNumber);
			}
			
			if(row.getCell(columnNumber) == null)
			{
				cell = row.createCell(columnNumber);
			}
			else
			{
				cell = row.getCell(columnNumber);
			}
			
			
			cell.setCellType(CellType.STRING);
			XSSFCellStyle cellStyle = workbook.createCellStyle();
			
			applyTestResult(cell, cellStyle, testCaseResult);
			
			cellStyle.setWrapText(true);
			cell.setCellStyle(cellStyle);
			
			currentSheet.autoSizeColumn(columnNumber);
			row.setHeight((short)-1);

			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void markTestCaseResults(ConcurrentHashMap<String, TestCaseResult> testResults)
	{
		try
		{
			if(!testResults.isEmpty())
			{
				openFile(filePath, currentSheet.getSheetName());
				
				Iterator<String> mapIterator = testResults.keySet().iterator();
				while(mapIterator.hasNext())
				{
					String			currentKey		= mapIterator.next();
					CellReference	ref				= new CellReference(currentKey);
					int				rowNumber		= ref.getRow();
					int				columnNumber	= ref.getCol();
					
					XSSFRow  row;
					XSSFCell cell;
					
					if(currentSheet.getRow(rowNumber) == null)
					{
						row  = currentSheet.createRow(rowNumber);
					}
					else
					{
						row  = currentSheet.getRow(rowNumber);
					}
					
					if(row.getCell(columnNumber) == null)
					{
						cell = row.createCell(columnNumber);
					}
					else
					{
						cell = row.getCell(columnNumber);
					}
					
					
					cell.setCellType(CellType.STRING);
					XSSFCellStyle cellStyle = workbook.createCellStyle();
					
					applyTestResult(cell, cellStyle, testResults.get(currentKey));
					
					cellStyle.setWrapText(true);
					cell.setCellStyle(cellStyle);
					
					currentSheet.autoSizeColumn(columnNumber);
					if(currentSheet.getColumnWidth(columnNumber) < 2350)
					{
						currentSheet.setColumnWidth(columnNumber, 2350);
					}
					row.setHeight((short)-1);
				}
	
				FileOutputStream fos = new FileOutputStream(filePath);
				workbook.write(fos);
				fos.close();
			}
		}
		catch(Exception e)
		{
			
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void markTestCaseResultByName(String testCaseName, TestCaseResult testCaseResult)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			Iterator<Row>							rows				= currentSheet.iterator();																									// a spreadsheet is a collection of rows
			Row										firstRow			= rows.next();
			Iterator<Cell>							cells				= firstRow.cellIterator();																									// a row is a collection of cells
			int										x					= firstRow.getFirstCellNum();
			int										nameColumnNumber	= -1;
			int										resultColumnNumber	= -1;
			SimpleEntry<String, String>				keyMap				= columnCache.get(new SimpleEntry<String, String>(new File(filePath).getName(), currentSheet.getSheetName()));
			String									nameColumnHeader	= keyMap.getKey();
			String									resultColumnHeader	= keyMap.getValue();

			while(cells.hasNext())
			{
				Cell	currentCell		= cells.next();
				String	cellData		= formatter.formatCellValue(currentCell);
	
				if(cellData.equalsIgnoreCase(nameColumnHeader))
				{
					nameColumnNumber	= x;
				}
				
				if(cellData.equalsIgnoreCase(resultColumnHeader))
				{
					resultColumnNumber	= x;
				}
				
				if((nameColumnNumber != -1) && (resultColumnNumber != -1))
				{
					break;
				}
	
				x++;
			}
	
			while(rows.hasNext())
			{
				Row			currentRow	= rows.next();
				XSSFCell	nameCell	= (XSSFCell)(currentRow.getCell(nameColumnNumber)   == null ? currentRow.createCell(nameColumnNumber)   : currentRow.getCell(nameColumnNumber));
				XSSFCell	resultCell	= (XSSFCell)(currentRow.getCell(resultColumnNumber) == null ? currentRow.createCell(resultColumnNumber) : currentRow.getCell(resultColumnNumber));
				String		cellData	= formatter.formatCellValue(nameCell);
	
				if(cellData.equalsIgnoreCase(testCaseName))
				{
					resultCell.setCellType(CellType.STRING);
					XSSFCellStyle cellStyle = workbook.createCellStyle();
					
					applyTestResult(resultCell, cellStyle, testCaseResult);
					
					cellStyle.setWrapText(true);
					resultCell.setCellStyle(cellStyle);
					
					break;
				}
				
				currentSheet.autoSizeColumn(nameColumnNumber);
				if(currentSheet.getColumnWidth(nameColumnNumber) < 2350)
				{
					currentSheet.setColumnWidth(nameColumnNumber, 2350);
				}
				
				currentSheet.autoSizeColumn(resultColumnNumber);
				if(currentSheet.getColumnWidth(resultColumnNumber) < 2350)
				{
					currentSheet.setColumnWidth(resultColumnNumber, 2350);
				}
				
				currentRow.setHeight((short)-1);
			}
			
			keyMap = null;
			
			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void markTestCaseResultsByName(ConcurrentHashMap<String, TestCaseResult> testResults)
	{
		try
		{
			if(!testResults.isEmpty())
			{
				openFile(filePath, currentSheet.getSheetName());
				
				Iterator<String> mapIterator = testResults.keySet().iterator();
				while(mapIterator.hasNext())
				{
					Iterator<Row>				rows				= currentSheet.iterator();			// a spreadsheet is a collection of rows
					Row							firstRow			= rows.next();
					Iterator<Cell>				cells				= firstRow.cellIterator();			// a row is a collection of cells
					int							x					= firstRow.getFirstCellNum();
					int							nameColumnNumber	= -1;
					int							resultColumnNumber	= -1;
					SimpleEntry<String, String>	keyMap				= columnCache.get(new SimpleEntry<String, String>(new File(filePath).getName(), currentSheet.getSheetName()));
					String						nameColumnHeader	= keyMap.getKey();
					String						resultColumnHeader	= keyMap.getValue();
					String						testCaseName		= mapIterator.next();
					TestCaseResult				testCaseResult		= testResults.get(testCaseName);
					
					while(cells.hasNext())
					{
						Cell	currentCell		= cells.next();
						String	cellData		= formatter.formatCellValue(currentCell);
			
						if(cellData.equalsIgnoreCase(nameColumnHeader))
						{
							nameColumnNumber	= x;
						}
						
						if(cellData.equalsIgnoreCase(resultColumnHeader))
						{
							resultColumnNumber	= x;
						}
						
						if((nameColumnNumber != -1) && (resultColumnNumber != -1))
						{
							break;
						}
			
						x++;
					}
					
					while(rows.hasNext())
					{
						Row			currentRow	= rows.next();
						XSSFCell	nameCell	= (XSSFCell)(currentRow.getCell(nameColumnNumber)   == null ? currentRow.createCell(nameColumnNumber)   : currentRow.getCell(nameColumnNumber));
						XSSFCell	resultCell	= (XSSFCell)(currentRow.getCell(resultColumnNumber) == null ? currentRow.createCell(resultColumnNumber) : currentRow.getCell(resultColumnNumber));
						String		cellData	= formatter.formatCellValue(nameCell);
			
						if(cellData.equalsIgnoreCase(testCaseName))
						{
							resultCell.setCellType(CellType.STRING);
							XSSFCellStyle cellStyle = workbook.createCellStyle();
							
							applyTestResult(resultCell, cellStyle, testCaseResult);
							
							cellStyle.setWrapText(true);
							resultCell.setCellStyle(cellStyle);
							
							break;
						}
						
						currentSheet.autoSizeColumn(nameColumnNumber);
						if(currentSheet.getColumnWidth(nameColumnNumber) < 2350)
						{
							currentSheet.setColumnWidth(nameColumnNumber, 2350);
						}
						
						currentSheet.autoSizeColumn(resultColumnNumber);
						if(currentSheet.getColumnWidth(resultColumnNumber) < 2350)
						{
							currentSheet.setColumnWidth(resultColumnNumber, 2350);
						}
						
						currentRow.setHeight((short)-1);
					}
					
					keyMap = null;
				}
	
				FileOutputStream fos = new FileOutputStream(filePath);
				workbook.write(fos);
				fos.close();
			}
		}
		catch(Exception e)
		{
			
		}
		finally
		{
			closeFile();
		}
	}
	
	
	
	
	//	Pass/Fail/Skipped marker formats
	private void applyTestResult(XSSFCell cell, XSSFCellStyle cellStyle, TestCaseResult testCaseResult)
	{
		CellStyleHelper	csHelper = new CellStyleHelper(workbook, cellStyle);
		
		cell.setCellValue(testCaseResult.toString());
		
		if(testCaseResult == TestCaseResult.PASS)
		{
			csHelper.setFont("Calibri", (short)11, IndexedColors.GREEN.getIndex(), true, FontUnderline.NONE, true);
			csHelper.setAllBorderStyles(BorderStyle.MEDIUM, limeGreen, true, true, true, true);
			csHelper.setFill(limeGreen, FillPatternType.DIAMONDS);
		}
		else
			if(testCaseResult == TestCaseResult.FAIL)
			{
				csHelper.setFont("Calibri", (short)11, IndexedColors.BROWN.getIndex(), true, FontUnderline.NONE, true);
				csHelper.setAllBorderStyles(BorderStyle.MEDIUM, IndexedColors.RED.getIndex(), true, true, true, true);
				csHelper.setFill(IndexedColors.RED.getIndex(), FillPatternType.DIAMONDS);
			}
		else
			if(testCaseResult == TestCaseResult.SKIPPED)
			{
				csHelper.setFont("Calibri", (short)11, IndexedColors.GOLD.getIndex(), true, FontUnderline.NONE, true);
				csHelper.setAllBorderStyles(BorderStyle.MEDIUM, IndexedColors.GOLD.getIndex(), true, true, true, true);
				csHelper.setFill(IndexedColors.GOLD.getIndex(), FillPatternType.DIAMONDS);
			}
	}
	
	
	
	
	//	Check if test should be ran
	public void checkRunModeForSuite(String suiteListSheetName, String suiteName, String testCaseName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			
			//	suite level
			XSSFSheet suitesSheet  = getSheet(suiteListSheetName);
			
			if(suitesSheet == null)
			{
				return;
			}
			
			
			//	test level
			XSSFSheet testsSheet  = getSheet(suiteName);
			
			if(testsSheet == null)
			{
				return;
			}
			
			
			if(!isSuiteRunnable(suitesSheet, suiteName))
			{
				throw new SkipException("Skipping test case: " + testCaseName + ", the RunMode for suite: " + suiteName + " is set to SKIP.");
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void checkRunModeForTest(String suiteListSheetName, String suiteName, String testName, String testCaseName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			
			//	suite level
			XSSFSheet suitesSheet  = getSheet(suiteListSheetName);
			
			if(suitesSheet == null)
			{
				return;
			}
			
			
			//	test level
			XSSFSheet testsSheet  = getSheet(suiteName);
			
			if(testsSheet == null)
			{
				return;
			}
			
			
			if(!isTestRunnable(suitesSheet, testsSheet, testName))
			{
				throw new SkipException("Skipping test case: " + testCaseName + ", the Runmode for test: " + testName + " is set to SKIP.");
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	public void checkRunModeForTestCase(String suiteListSheetName, String suiteName, String testName, String testCaseName)
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			
			//	suite level
			XSSFSheet suitesSheet  = getSheet(suiteListSheetName);
			
			if(suitesSheet == null)
			{
				return;
			}
			
			
			//	test level
			XSSFSheet testsSheet  = getSheet(suiteName);
			
			if(testsSheet == null)
			{
				return;
			}
			
			
			//	testcase level
			XSSFSheet testcasesSheet  = getSheet(testName);
			
			if(testcasesSheet == null)
			{
				return;
			}
			
			
			if(!isTestCaseRunnable(suitesSheet, testsSheet, testcasesSheet, testName, testCaseName))
			{
				throw new SkipException("Skipping test case: " + testCaseName + ", its run mode is set to SKIP.");
			}
		}
		catch(Exception e)
		{
			onErrorDo(e);
		}
		finally
		{
			closeFile();
		}
	}
	
	
	private XSSFSheet getSheet(String sheetName)
	{
		int sheetIndex = workbook.getSheetIndex(sheetName);

		if(sheetIndex == -1)
		{
			System.err.println("Unable to open sheet: " + sheetName + ".\r\nCurrent workbook does not have a sheet by this name.");
			return null;
		}
		
		return workbook.getSheetAt(sheetIndex);
	}


	private boolean isSuiteRunnable(XSSFSheet suiteListSheet, String suiteName)
	{
		int rows = getRowCount(suiteListSheet.getSheetName(), false);
		
		if(rows == 0)
		{
			return false;
		}
		
		int	firstRowNum		= suiteListSheet.getFirstRowNum();
		int	suiteNameColumn	= -1;
		int	runModeColumn	= -1;
		Row	firstRow		= suiteListSheet.getRow(firstRowNum);
		
		for(int colNum = firstRow.getFirstCellNum(); colNum <= firstRow.getLastCellNum(); colNum++)
		{
			Cell cell = firstRow.getCell(colNum);
			
			if(cell == null)
			{
				continue;
			}
			
			String data = getCellData(colNum + 1, firstRowNum + 1);
			
			if(data.equalsIgnoreCase("Suite"))
			{
				suiteNameColumn = colNum;
			}
			else
				if(data.equalsIgnoreCase("Run Mode"))
				{
					runModeColumn = colNum;
				}
			
			if((suiteNameColumn != -1) && (runModeColumn != -1))
			{
				break;
			}
		}

		for(int rowNum = firstRowNum + 1; rowNum <= rows; rowNum++)
		{
			Row		row				= suiteListSheet.getRow(rowNum);
			Cell	suiteNameCell	= row.getCell(suiteNameColumn);
			
			if(!formatter.formatCellValue(suiteNameCell).equalsIgnoreCase(suiteName))
			{
				continue;
			}
			
			Cell	runModeCell		= row.getCell(runModeColumn);
			
			if(formatter.formatCellValue(runModeCell).equalsIgnoreCase(RunMode.RUN.toString()))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		return false;
	}


	private boolean isTestRunnable(XSSFSheet suiteListSheet, XSSFSheet testListSheet, String testName)
	{
		if(isSuiteRunnable(suiteListSheet, testListSheet.getSheetName()) == false)
		{
			return false;
		}
		
		int rows = getRowCount(testListSheet.getSheetName(), false);
		
		if(rows == 0)
		{
			return false;
		}
		
		int	firstRowNum		= testListSheet.getFirstRowNum();
		int	testNameColumn	= -1;
		int	runModeColumn	= -1;
		Row	firstRow		= testListSheet.getRow(firstRowNum);
		
		for(int colNum = firstRow.getFirstCellNum(); colNum <= firstRow.getLastCellNum(); colNum++)
		{
			Cell cell = firstRow.getCell(colNum);
			
			if(cell == null)
			{
				continue;
			}
			
			String data = getCellData(colNum + 1, firstRowNum + 1);
			
			if(data.equalsIgnoreCase("Test"))
			{
				testNameColumn = colNum;
			}
			else
				if(data.equalsIgnoreCase("Run Mode"))
				{
					runModeColumn  = colNum;
				}
			
			if((testNameColumn != -1) && (runModeColumn != -1))
			{
				break;
			}
		}
		
		for(int rowNum = firstRowNum + 1; rowNum <= rows; rowNum++)
		{
			Row		row				= testListSheet.getRow(rowNum);
			Cell	testNameCell	= row.getCell(testNameColumn);
			
			if(!formatter.formatCellValue(testNameCell).equalsIgnoreCase(testName))
			{
				continue;
			}
			
			Cell	runModeCell		= row.getCell(runModeColumn);
			
			if(formatter.formatCellValue(runModeCell).equalsIgnoreCase(RunMode.RUN.toString()))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		return false;
	}
	
	
	private boolean isTestCaseRunnable(XSSFSheet suiteListSheet, XSSFSheet testListSheet, XSSFSheet testcaseListSheet, String testName, String testcaseName)
	{
		if(isTestRunnable(suiteListSheet, testListSheet, testName) == false)
		{
			return false;
		}
		
		int rows = getRowCount(testcaseListSheet.getSheetName(), false);
		
		if(rows == 0)
		{
			return false;
		}
		
		int	firstRowNum			= testcaseListSheet.getFirstRowNum();
		int	testcaseNameColumn	= -1;
		int	runModeColumn		= -1;
		Row	firstRow			= testcaseListSheet.getRow(firstRowNum);
		
		for(int colNum = firstRow.getFirstCellNum(); colNum <= firstRow.getLastCellNum(); colNum++)
		{
			Cell cell = firstRow.getCell(colNum);
			
			if(cell == null)
			{
				continue;
			}
			
			String data = getCellData(colNum + 1, firstRowNum + 1);
			
			if(data.equalsIgnoreCase("Testcase"))
			{
				testcaseNameColumn	= colNum;
			}
			else
				if(data.equalsIgnoreCase("Run Mode"))
				{
					runModeColumn  		= colNum;
				}
			
			if((testcaseNameColumn != -1) && (runModeColumn != -1))
			{
				break;
			}
		}
		
		for(int rowNum = firstRowNum + 1; rowNum <= rows; rowNum++)
		{
			Row		row					= testcaseListSheet.getRow(rowNum);
			Cell	testcaseNameCell	= row.getCell(testcaseNameColumn);
			
			if(!formatter.formatCellValue(testcaseNameCell).equalsIgnoreCase(testcaseName))
			{
				continue;
			}
			
			Cell	runModeCell			= row.getCell(runModeColumn);
			
			if(formatter.formatCellValue(runModeCell).equalsIgnoreCase(RunMode.RUN.toString()))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		return false;
	}
	
	
	
	
	//	Playing around with predefined colors
	public void testHSSFColors(String cellReference) throws IOException
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			CellReference	ref				= new CellReference(cellReference);
			int				rowNumber		= ref.getRow();
			int				columnNumber	= ref.getCol();
			
			for(HSSFColor.HSSFColorPredefined color : HSSFColor.HSSFColorPredefined.values())
			{
				XSSFRow		row;
				XSSFCell	cell;
				
				if(currentSheet.getRow(rowNumber) == null)
				{
					row = currentSheet.createRow(rowNumber);
				}
				else
				{
					row = currentSheet.getRow(rowNumber);
				}
				
				if(row.getCell(columnNumber) == null)
				{
					cell = row.createCell(columnNumber);
				}
				else
				{
					cell = row.getCell(columnNumber);
				}
				
				
				cell.setCellValue(color.toString() + " - " + color.getIndex());
				
				XSSFCellStyle	cellStyle	= workbook.createCellStyle();
				CellStyleHelper	csHelper	= new CellStyleHelper(workbook, cellStyle);
				
				csHelper.setFont("Calibri", (short)11, color.getIndex(), true, FontUnderline.NONE, true);
				csHelper.setAllBorderStyles(BorderStyle.MEDIUM, color.getIndex(), true, true, true, true);
				csHelper.setFill(color.getIndex(), FillPatternType.LESS_DOTS);
				cell.setCellStyle(cellStyle);
				
				currentSheet.autoSizeColumn(columnNumber);
				row.setHeight((short)-1);
				
				rowNumber++;
			}

			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
		}
		finally
		{
			closeFile();
		}
	}


	public void testXSSFColors(String cellReference) throws IOException
	{
		try
		{
			openFile(filePath, currentSheet.getSheetName());
			
			CellReference	ref				= new CellReference(cellReference);
			int				rowNumber		= ref.getRow();
			int				columnNumber	= ref.getCol();
			
			for(IndexedColors ic : IndexedColors.values())
			{
				XSSFRow		row;
				XSSFCell	cell;
				
				if(currentSheet.getRow(rowNumber) == null)
				{
					row = currentSheet.createRow(rowNumber);
				}
				else
				{
					row = currentSheet.getRow(rowNumber);
				}
				
				if(row.getCell(columnNumber) == null)
				{
					cell = row.createCell(columnNumber);
				}
				else
				{
					cell = row.getCell(columnNumber);
				}
				
				
				cell.setCellValue(ic.toString() + " - " + ic.getIndex());
				
				XSSFCellStyle	cellStyle	= workbook.createCellStyle();
				CellStyleHelper	csHelper	= new CellStyleHelper(workbook, cellStyle);
				
				csHelper.setFont("Calibri", (short)11, ic.getIndex(), true, FontUnderline.NONE, true);
				csHelper.setAllBorderStyles(BorderStyle.MEDIUM, ic.getIndex(), true, true, true, true);
				csHelper.setFill(ic.getIndex(), FillPatternType.LESS_DOTS);
				cell.setCellStyle(cellStyle);
				
				currentSheet.autoSizeColumn(columnNumber);
				row.setHeight((short)-1);
				
				rowNumber++;
			}

			FileOutputStream fos = new FileOutputStream(filePath);
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
		}
		finally
		{
			closeFile();
		}
	}
}