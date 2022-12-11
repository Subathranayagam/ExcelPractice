package org.excelpracticeone;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class WriteExcelNew {
	@Test
	public void writeExcelNew() throws IOException
	{
		File file = new File(System.getProperty("user.dir") +"/src/test/resources/Book1.xlsx");
		FileInputStream input = new FileInputStream(file);
		
		XSSFWorkbook workbook  = new XSSFWorkbook(input);
		XSSFSheet createSheet = workbook.createSheet("Subathrass");
		XSSFSheet sheet = workbook.getSheet("Subathrass");
		sheet.createRow(0).createCell(0).setCellValue("Haihello");
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		workbook.close();
		out.close();
		
	}

}
