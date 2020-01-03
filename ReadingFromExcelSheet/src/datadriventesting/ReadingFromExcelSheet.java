package datadriventesting;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;//if we put * will not ask to importimport org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook.*;
public class ReadingFromExcelSheet {

	public static void main(String[] args) throws IOException {
//The Sequence for reading data from Excel is 
//Excel--Workbooks--sheets--rows--cells
     
		//this will opens the Excel input/read mode
		FileInputStream File = new FileInputStream("‪https://onedrive.live.com/edit.aspx?resid=7A6DF7E9C0921A1E!116&ithint=file%2cxlsx");//
		
		XSSFWorkbook workbook = new XSSFWorkbook(File);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");//providing with sheet Name Method
		//XSSFSheet sheet = workbook.getSheetAt(0);//providing index Method
		
		//this will give us the last row number/number of records
		int rowcount = sheet.getLastRowNum();//this will return row count
		
		int colomcount = sheet.getRow(0).getLastCellNum();//this will return the coloum/cell count
		
		for(int i=0;i<rowcount;i++)//increment the row count
		{
			XSSFRow currentrow = sheet.getRow(i);//focusing on current row
			
			for(int j = 0; j<colomcount;j++) //this will increment current row
			{
			String value = currentrow.getCell(j).toString();//reads the data from the cell
			System.out.print("  " +value);
		}
			System.out.println();
	}

}
}