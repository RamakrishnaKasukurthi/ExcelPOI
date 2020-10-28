package com.excel.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	public void Xls_Reader(String path) {

		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	public int getRowCount(String sheetName) {
		
		sheet=workbook.getSheet(sheetName);
		int number=sheet.getLastRowNum();
		return number;
		
		
	}
	
	/*
	 * public static String getCellData(String sheetName, int rowNum, int cellNum) {
	 * sheet=workbook.getSheet(sheetName);
	 * cell=sheet.getRow(rowNum).getCell(cellNum);
	 * 
	 * String cellData=cell.getStringCellValue(); return cellData;
	 * 
	 * }
	 */

	/*
	 * 
	 * // Write data in the excel. FileOutputStream foutput=new FileOutputStream(src);
	 * 
	 * // Specify the message needs to be written. 
	 * String message = "Data Imported Successfully.";
	 * 
	 * // Create cell where data needs to be written.
	 * sheet.getRow(i).createCell(3).setCellValue(message);
	 * 
	 * // Specify the file in which data needs to be written. FileOutputStream
	 * fileOutput = new FileOutputStream(src);
	 * 
	 * // finally write content 
	 * workbook.write(fileOutput);
	 * 
	 * // close the file 
	 * fileOutput.close();
	 */	
}

