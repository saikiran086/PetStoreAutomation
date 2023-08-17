package api.utilities;

import java.io.IOException;

public class Dummy {

	public static void main(String[] args) throws IOException {
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
		
		System.out.println(apidata);
		
	}

}
