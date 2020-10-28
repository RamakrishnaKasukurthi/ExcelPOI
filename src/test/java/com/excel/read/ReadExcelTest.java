package com.excel.read;

import org.testng.annotations.Test;

import com.excel.util.ExcelUtils;

public class ReadExcelTest extends ExcelUtils{
	//ExcelUtils excelUtils;
	
	public String excelPath="C:\\Users\\91800\\eclipse-workspace\\ExcelPOI\\src\\main\\java\\com\\excel\\TestData\\TestData.xlsx";
	public String excelSheetName="Info";
	
	
	@Test
	public void readExcelData(){
		
	ExcelUtils excelUtils=new ExcelUtils();
	excelUtils.Xls_Reader(excelPath);
	
		
		
	}

}
