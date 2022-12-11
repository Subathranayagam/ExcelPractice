package org.excelpracticeone;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class WriteExcel {
	
	@Test
	public void writeExcel() throws IOException
	{
		File file = new File(System.getProperty("user.dir")+"/src/test/resources/Book1.xlsx");
		FileInputStream input = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int totalRows = sheet.getPhysicalNumberOfRows();
		XSSFRow createRow = sheet.createRow(totalRows);
		for(int i=0;i<3;i++)
		{
			createRow.createCell(i).setCellValue("Subathra");
		}
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		workbook.close();
		out.close();	
	}

}
