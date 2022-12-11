package org.excelpractice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelPractice {
	
@Test
public void excelDemo() throws IOException
{
	
	File file = new File(System.getProperty("user.dir") +"/src/test/resources/Students details - Nov project 4.xlsx");
	FileInputStream input = new FileInputStream(file);
	XSSFWorkbook workbook = new XSSFWorkbook(input);
	XSSFSheet sheet = workbook.getSheet("Vishnu Mar Batch");
	int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
	XSSFRow createRow = sheet.createRow(physicalNumberOfRows);
	createRow.createCell(0).setCellValue("Subathra");
	createRow.createCell(1).setCellValue("Subathra");
	createRow.createCell(2).setCellValue("Subathra");
	FileOutputStream output = new FileOutputStream(file);
	workbook.write(output);
	workbook.close();
	output.close();
}
}