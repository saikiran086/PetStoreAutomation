package api.utilities;

import java.io.IOException;

import org.testng.annotations.DataProvider;

public class DataProviders 
{
	@DataProvider(name="Data")
	public String[][] getAllData() throws Exception 
	{
		
		ExcelUtility eu=new ExcelUtility("C:\\Users\\dulam\\OneDrive\\Desktop\\userdata.xlsx");
				
		int rownum=eu.getRows("Sheet1");
		int colCount=eu.getCells("Sheet1", 1);
		String apidata[][]=new String [rownum][colCount];
		
		for(int i=1; i<=rownum; i++)
		{
			for(int j=0; j<colCount; j++)
			{
				apidata[i-1][j]=eu.getData("Sheet1", i, j);
			}

		}
		return apidata;
	}

	@DataProvider(name="userNames")
	public String[] getUserNames() throws Exception
	{
		ExcelUtility eu=new ExcelUtility("C:\\Users\\dulam\\OneDrive\\Desktop\\userdata.xlsx");
		
		
		int rownum=eu.getRows("Sheet1");

		String apidata[] = new String[rownum];
		
		for(int i=1; i<=rownum; i++)
		{
			apidata[i-1]=eu.getData("Sheet1", i, 1);
		}
		return apidata;
	}
}




























