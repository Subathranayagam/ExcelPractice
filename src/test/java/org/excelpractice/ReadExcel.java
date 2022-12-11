package org.excelpractice;

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

public class ReadExcel {
	
	@Test
	public void readExcel() throws IOException
	{
		File file = new File(System.getProperty("user.dir") +"/src/test/resources/Book1.xlsx");
		FileInputStream input = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
     	int rows = sheet.getLastRowNum();
     	int col = sheet.getRow(1).getLastCellNum();
     	for(int i=0;i<=rows;i++)
     	{
     		XSSFRow row = sheet.getRow(i);
     		for(int j=0;j<col;j++)
     		{
     			XSSFCell cell = row.getCell(j);
     			switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;

				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;

				}
     			System.out.print(" | ");
     		}
     		System.out.println();
     	}

}
}
