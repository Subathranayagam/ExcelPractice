package org.excelpracticeone;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.compress.utils.FixedLengthBlockOutputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ModifyExcel {
	@Test
	public void modifyExcel() throws IOException
	{
		
	File file = new File(System.getProperty("user.dir") +"/src/test/resources/Book1.xlsx");
	FileInputStream input = new FileInputStream(file);
	XSSFWorkbook workbook = new XSSFWorkbook(input);
	XSSFSheet sheet = workbook.getSheet("Sheet1");
	XSSFRow row = sheet.getRow(0);
	row.getCell(0).setCellValue("Delhi");
	FileOutputStream outputStream = new FileOutputStream(file);
	workbook.write(outputStream);
	outputStream.close();
	
	
	}

}
