package org.excelpracticeone;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class NewExcelCreate {
	@Test
	public void newExcelCreate() throws IOException
	{
		File file = new File("C:\\Users\\Subathra\\eclipse-workspace1\\ExcelPractice\\src\\test\\resources\\Book2.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet createSheet = workbook.createSheet("Subathra");
		XSSFSheet sheet = workbook.getSheet("Subathra");
		sheet.createRow(0).createCell(0).setCellValue("Haihello");
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		workbook.close();
		out.close();
	}
	
	
	}


