package excelNdp;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Integration1 {
		
	DataFormatter df =new DataFormatter();
		@Test(dataProvider="getData")
		public void testCase(String d1,String d2,String d3,String d4) {
			System.out.println(d1 + d2 + d3 + d4);
		}
		
		@DataProvider
		public Object[][] getData() throws IOException {
			
			FileInputStream fis =new FileInputStream("C:\\Users\\aadhi\\OneDrive\\BookDemo.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);
			int rCount =sheet.getPhysicalNumberOfRows();
			XSSFRow row =sheet.getRow(0);
			int cCount =row.getLastCellNum();
			Object dataset[][] = new Object[rCount-1][cCount];
			for(int i=0;i<rCount-1;i++)
			{
				row= sheet.getRow(i+1);
				for(int j=0;j<cCount;j++)
				{
					XSSFCell rc =row.getCell(j);
					dataset[i][j] =df.formatCellValue(rc);	
				}
			}
			return dataset;
		}
		
		
		
		
}
