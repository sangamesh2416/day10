package com.selenium.Cse1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;

import org.apache.commons.math3.analysis.function.Max;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class D10_3 {
	
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
					 ArrayList<Double> arrayList=new ArrayList<Double>();
					double s=0;
		            for (Cell cell : row) { // moves each cell in the particular row
		                CellType cellType = cell.getCellType();// returns kind of data in cell
		                if (cellType == CellType.STRING) {
		                    //System.out.print(cell.getStringCellValue() + "\t"); //get string values from cell
		                } else if (cellType == CellType.NUMERIC) {
		                   // System.out.print(cell.getNumericCellValue() + "\t");//get numerical values from cell
		                	arrayList.add(cell.getNumericCellValue());
		                }
		            }
		            
		            System.out.println(Collections.min(arrayList));
		        }
			}

}
