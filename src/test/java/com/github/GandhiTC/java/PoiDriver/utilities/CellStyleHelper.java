package com.github.GandhiTC.java.PoiDriver.utilities;



import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class CellStyleHelper
{
	private XSSFWorkbook	workbook;
	private XSSFCellStyle	cellStyle;


	
	
	// Constructors
	public CellStyleHelper(XSSFWorkbook workbook, XSSFCellStyle cellStyle)
	{
		this.workbook	= workbook;
		this.cellStyle	= cellStyle;
	}
	
	
	public CellStyleHelper(XSSFWorkbook workbook)
	{
		this.workbook	= workbook;
		this.cellStyle	= this.workbook.createCellStyle();
	}
	
	
	
	
	// Getter for cellStyle
	public XSSFCellStyle cellStyle()
	{
		return this.cellStyle;
	}
	
	
	
	
	// Point current instance of CellStyleHelper to a different Workbook and/or CellStyle
	public void switchTo(XSSFWorkbook workbook, XSSFCellStyle cellStyle)
	{
		this.workbook	= workbook;
		this.cellStyle	= cellStyle;
	}
	
	
	
	
	// Set cell font
	public void setFont(String fontName, short fontSize,
			XSSFColor fontColor, boolean fontBold, FontUnderline fontUnderline, boolean fontItalic)
	{
		XSSFFont font = workbook.createFont();
		font.setFontName(fontName);
		font.setFontHeightInPoints(fontSize);
		font.setColor(fontColor);
		font.setBold(fontBold);
		font.setUnderline(fontUnderline);
		font.setItalic(fontItalic);
		cellStyle.setFont(font);
	}


	public void setFont(String fontName, short fontSize,
			short fontColor, boolean fontBold, FontUnderline fontUnderline, boolean fontItalic)
	{
		XSSFFont font = workbook.createFont();
		font.setFontName(fontName);
		font.setFontHeightInPoints(fontSize);
		font.setColor(fontColor);
		font.setBold(fontBold);
		font.setUnderline(fontUnderline);
		font.setItalic(fontItalic);
		cellStyle.setFont(font);
	}

	
	

	// Set cell border styles and colors
	public void setAllBorderStyles(BorderStyle borderStyle, short borderColor,
			boolean applyToTop, boolean applyToBottom, boolean applyToLeft, boolean applyToRight)
	{

		if(applyToTop)
		{
			cellStyle.setBorderTop(borderStyle);
			cellStyle.setTopBorderColor(borderColor);
		}

		if(applyToBottom)
		{
			cellStyle.setBorderBottom(borderStyle);
			cellStyle.setBottomBorderColor(borderColor);
		}

		if(applyToLeft)
		{
			cellStyle.setBorderLeft(borderStyle);
			cellStyle.setLeftBorderColor(borderColor);
		}

		if(applyToRight)
		{
			cellStyle.setBorderRight(borderStyle);
			cellStyle.setRightBorderColor(borderColor);
		}

	}


	public void setAllBorderStyles(BorderStyle borderStyle, boolean applyToTop,
			boolean applyToBottom, boolean applyToLeft, boolean applyToRight)
	{

		if(applyToTop)
		{
			cellStyle.setBorderTop(borderStyle);
		}

		if(applyToBottom)
		{
			cellStyle.setBorderBottom(borderStyle);
		}

		if(applyToLeft)
		{
			cellStyle.setBorderLeft(borderStyle);
		}

		if(applyToRight)
		{
			cellStyle.setBorderRight(borderStyle);
		}

	}


	public void setAllBorderStyles(short borderColor, boolean applyToTop,
			boolean applyToBottom, boolean applyToLeft, boolean applyToRight)
	{

		if(applyToTop)
		{
			cellStyle.setTopBorderColor(borderColor);
		}

		if(applyToBottom)
		{
			cellStyle.setBottomBorderColor(borderColor);
		}

		if(applyToLeft)
		{
			cellStyle.setLeftBorderColor(borderColor);
		}

		if(applyToRight)
		{
			cellStyle.setRightBorderColor(borderColor);
		}

	}


	public void setAllBorderStyles(BorderStyle borderStyle, XSSFColor borderColor,
			boolean applyToTop, boolean applyToBottom, boolean applyToLeft, boolean applyToRight)
	{

		if(applyToTop)
		{
			cellStyle.setBorderTop(borderStyle);
			cellStyle.setTopBorderColor(borderColor);
		}

		if(applyToBottom)
		{
			cellStyle.setBorderBottom(borderStyle);
			cellStyle.setBottomBorderColor(borderColor);
		}

		if(applyToLeft)
		{
			cellStyle.setBorderLeft(borderStyle);
			cellStyle.setLeftBorderColor(borderColor);
		}

		if(applyToRight)
		{
			cellStyle.setBorderRight(borderStyle);
			cellStyle.setRightBorderColor(borderColor);
		}

	}


	public void setAllBorderStyles(XSSFColor borderColor, boolean applyToTop,
			boolean applyToBottom, boolean applyToLeft, boolean applyToRight)
	{

		if(applyToTop)
		{
			cellStyle.setTopBorderColor(borderColor);
		}

		if(applyToBottom)
		{
			cellStyle.setBottomBorderColor(borderColor);
		}

		if(applyToLeft)
		{
			cellStyle.setLeftBorderColor(borderColor);
		}

		if(applyToRight)
		{
			cellStyle.setRightBorderColor(borderColor);
		}

	}

	
	

	// Set cell fill color and style
	public void setFill(short bgColor, FillPatternType bgFillPatternType)
	{
		cellStyle.setFillForegroundColor(bgColor);
		cellStyle.setFillPattern(bgFillPatternType);
	}


	public void setFill(XSSFColor bgColor, FillPatternType bgFillPatternType)
	{
		cellStyle.setFillForegroundColor(bgColor);
		cellStyle.setFillPattern(bgFillPatternType);
	}
	
	
	
	
	// Set wrap-text
	public void setWrapText(boolean wrapped)
	{
		cellStyle.setWrapText(wrapped);
	}
}