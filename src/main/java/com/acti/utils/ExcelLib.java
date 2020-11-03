package com.acti.utils;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelLib {
	
			
		static XSSFWorkbook wb;
		

		public ExcelLib()  {
			
			try
			{
				File file = new File("./src/test/resources/testdata/actidata.xlsx");
				FileInputStream fis = new FileInputStream(file);
				wb = new XSSFWorkbook(fis);
				
			}
			catch (Exception e)
			{
				System.out.println("excel file not found  "+e.getMessage());
			}
			
			
		}
		
		public int  getrowcount (int sheetnum)
		{
			return wb.getSheetAt(sheetnum).getLastRowNum();
		}
		
		public String getcelldata(int sheetnum, int row, int cell)
		{
			return wb.getSheetAt(sheetnum).getRow(row).getCell(cell).toString();
		}
	}



