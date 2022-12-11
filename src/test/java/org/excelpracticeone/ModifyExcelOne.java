package org.excelpracticeone;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ModifyExcelOne {
	@Test
	public void modifyExcelOne() throws IOException
	{
		File file = new File("C:\\Users\\Subathra\\eclipse-workspace1\\ExcelPractice\\src\\test\\resources\\Book1.xlsx");
		FileInputStream input = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Subathra");
		sheet.getRow(1).getCell(2).setCellValue(6000);
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();
		workbook.close();
	}

}
