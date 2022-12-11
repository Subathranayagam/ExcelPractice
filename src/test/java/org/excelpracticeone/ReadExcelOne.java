package org.excelpracticeone;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ReadExcelOne {
	
	@Test
	public void readExcelOne() throws IOException
	{
		File file = new File("C:\\Users\\Subathra\\eclipse-workspace1\\ExcelPractice\\src\\test\\resources\\Book1.xlsx");
		FileInputStream input = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Subathra");
		int lastRowNum = sheet.getLastRowNum();
		
		for(int i=0;i<lastRowNum;i++)
		{
			XSSFRow row = sheet.getRow(i);
			int lastCellNum = row.getLastCellNum();
			for(int j=0;j<lastCellNum;j++)
			{
				XSSFCell cell = row.getCell(j);
				if(cell.getCellType()==CellType.STRING)
				{
					System.out.print(cell.getStringCellValue());
				}
				else {
					System.out.print(cell.getNumericCellValue());
				}
				System.out.print(" | ");
			}
			System.out.println("");
		}
	
	}
	
	
	

}
