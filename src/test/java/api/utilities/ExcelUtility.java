package api.utilities;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {

	
	public FileInputStream fis;
	public XSSFWorkbook    wb;
	public XSSFSheet       sheet;
	public XSSFRow         row;
	public XSSFCell        cell;	
	
	
	String path;
	
	public ExcelUtility(String path) {
		
		this.path=path;
	}
	
	public int  getRows(String sheetname) throws IOException {
		
		fis = new FileInputStream(path);
		wb  = new XSSFWorkbook(fis);
		sheet = wb.getSheet(sheetname);
		
		int rows = sheet.getLastRowNum();
		
		wb.close();
		
		fis.close();
			
		return rows;
	}
	
	public int getCells(String sheetname,int rownum) throws IOException
	{
		fis = new FileInputStream(path);
		wb  = new XSSFWorkbook(fis);
		sheet = wb.getSheet(sheetname);
		row   = sheet.getRow(rownum);
		
		int cells = row.getLastCellNum();
		
	    wb.close();
		
		fis.close();
		
		
		return cells;
		
	}
	
	public String getData(String sheetname,int rownum,int cellnum) throws IOException {
		
		fis = new FileInputStream(path);
		wb  = new XSSFWorkbook(fis);
		sheet = wb.getSheet(sheetname);
		row   = sheet.getRow(rownum);
		cell  = row.getCell(cellnum);
		
		//String value = cell.getStringCellValue();
		
		DataFormatter formatter = new DataFormatter();
		String data;
		try
		{
			data = formatter.formatCellValue(cell);//Returnd the formatter value of a cell as a String 
		}
		catch(Exception e)
		{
			data="";
		}
		
		wb.close();
		
		fis.close();
		
		return data;
	}
	
	
}
