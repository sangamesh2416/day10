package com.selenium.Cse1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.Sheet;

public class D10_2 {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String excelFilePath = "D:\\data.xlsx";
		
		try {
			//creating file input object
			FileInputStream inputStream=new FileInputStream(excelFilePath);
			//creating workbook object
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			//select sheet using object
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			//call readdata
			readdata(sheet);
			
			//FileOutputStream outputStream=new FileOutputStream(excelFilePath);
			//workbook.write(outputStream);
				workbook.close();
	            inputStream.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static void readdata(XSSFSheet sheet)
	{ 
		for (Row row : sheet) { // for each row- moves row by row
			double s=0;
            for (Cell cell : row) { // moves each cell in the particular row
                CellType cellType = cell.getCellType();// returns kind of data in cell
                if (cellType == CellType.STRING) {
                    //System.out.print(cell.getStringCellValue() + "\t"); //get string values from cell
                } else if (cellType == CellType.NUMERIC) {
                   // System.out.print(cell.getNumericCellValue() + "\t");//get numerical values from cell
                	s=s+cell.getNumericCellValue();
                }
            }
            System.out.println(s);
        }
	}

}
