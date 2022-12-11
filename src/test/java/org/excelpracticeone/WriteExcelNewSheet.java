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

public class WriteExcelNewSheet {
	
	@Test
	public void writeExcelNewSheet() throws IOException
	{
		File file = new File("C:\\Users\\Subathra\\eclipse-workspace1\\ExcelPractice\\src\\test\\resources\\Book1.xlsx");
		FileInputStream input = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet createSheet = workbook.createSheet("Subathra3");
		XSSFSheet sheet = workbook.getSheet("Subathra3");
		XSSFRow createRow = sheet.createRow(0);
		createRow.createCell(0).setCellValue("Hai!");
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		//out.close();
		workbook.close();	
		
	}

}
