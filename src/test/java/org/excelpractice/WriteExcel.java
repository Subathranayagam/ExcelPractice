package org.excelpractice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class WriteExcel {
	@Test
	public void writeExcel() throws IOException, InvalidFormatException
	{
		File file = new File(System.getProperty("user.dir") +"/src/test/resources/Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		workbook.createSheet("Sheet1");
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		XSSFRow createRow = sheet.createRow(1);
		for(int i=0;i<10;i++)
		{
			createRow.createCell(i).setCellValue("Subathra");
		}
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		workbook.close();
		out.close();	
	}

}
