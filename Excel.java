--package com.application.test.utilities;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils
{
    public static XSSFSheet ExcelWSheet;
    public static XSSFWorkbook ExcelWBook;
    public static XSSFCell Cell;
    public static XSSFRow Row;
    
	public static void setExcelfile(String path , String sheetName) throws IOException
	{
	try
		{
		    FileInputStream FIS = new FileInputStream(path);
		    ExcelWBook = new XSSFWorkbook(FIS);
		    ExcelWSheet = ExcelWBook.getSheet(sheetName);
		}
		catch (Exception e)
		{
			throw (e);
		}
    //ExcelWSheet = ExcelWBook.getSheetAt(0);
	}
	
	public static String getStringCellData(int RowNum , int ColNum) throws IOException
	{
	    Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
	    String stringCellData = Cell.getStringCellValue();
	    return stringCellData;
	}
	
	public static Long getNumericCellData(int RowNum , int ColNum)throws IOException
	{
	    Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
	    Long numericCellData = (long)Cell.getNumericCellValue();
	    return numericCellData;
	}
	
	public static void setCellData(String path , String sheetName, String Result, int RowNum, int ColNum) throws Exception
	{
		try
		{
			Row  = ExcelWSheet.getRow(RowNum);
			Cell = Row.getCell(ColNum);
			if (Cell == null)
			{
				Cell = Row.createCell(ColNum);
				Cell.setCellValue(Result);
			}
			else
			{
				Cell.setCellValue(Result);
			}
			
			FileOutputStream FOS = new FileOutputStream(path);
			ExcelWBook.write(FOS);
			ExcelWSheet = ExcelWBook.getSheet(sheetName);
			FOS.flush();
		}
		catch(Exception e)
		{
			throw (e);
		}
	}
}
