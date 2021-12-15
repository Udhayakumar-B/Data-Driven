package com.Datadriven.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_Driven {
	public static void Particular_Data() throws IOException {
		File f = new File("C:\\Users\\User\\eclipse-Maven_NewProject\\Maven_Project\\DataDriven.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wp = new XSSFWorkbook(fis);
		Sheet sheetAt = wp.getSheetAt(0);
		Row row = sheetAt.getRow(0);
		Cell cell = row.getCell(1);
		
		CellType cellType = cell.getCellType();
		
		if (cellType.equals(cellType.STRING)) {
			String value = cell.getStringCellValue();
			System.out.println(value);
		}
		else if (cellType.equals(cellType.NUMERIC)) {
			double valuenum = cell.getNumericCellValue();
		long ref = (int) valuenum;
		System.out.println(ref);
		
		
		}
	
	}
public static void All_Data() throws IOException {
	
	File f1 = new File("C:\\Users\\User\\eclipse-Maven_NewProject\\Maven_Project\\DataDriven.xlsx");

	FileInputStream fi = new FileInputStream(f1);
	
	Workbook w = new XSSFWorkbook(fi);
	
	Sheet sheet = w.getSheetAt(0);
	
	int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
	
	for (int i = 0; i < physicalNumberOfRows; i++) {
		
		Row row = sheet.getRow(i);
		
		int NumberOfCells = row.getPhysicalNumberOfCells();
		
		for (int j = 0; j < NumberOfCells; j++) {
			
			Cell cell = row.getCell(j);
			CellType cellType = cell.getCellType();
			
			if(cellType.equals(CellType.STRING)) {
				
				String stringCellValue = cell.getStringCellValue();
				
				System.out.println(stringCellValue);
			}
			else if(cellType.equals(CellType.NUMERIC)) {
				
				double numericCellValue = cell.getNumericCellValue();
				int i1 = (int)numericCellValue;
				System.out.println(i1);
				
			}
		}
	}

	
	
}
	public static void Row_Data() throws IOException {
		File f1 = new File("C:\\Users\\User\\eclipse-Maven_NewProject\\Maven_Project\\DataDriven.xlsx");

		FileInputStream fis = new FileInputStream(f1);
		
		Workbook wp = new XSSFWorkbook(fis);
		
		Sheet sheetAt = wp.getSheetAt(0);
		
		Row row = sheetAt.getRow(4);
		int phNumberOfCells = row.getPhysicalNumberOfCells();
		
		for (int i = 0; i < phNumberOfCells; i++) {
			
			Cell c = row.getCell(i);
			CellType cType = c.getCellType();
			
			if(cType.equals(CellType.STRING)) {
				
				String sCellValue = c.getStringCellValue();
				
				System.out.println(sCellValue);
			}
			else if(cType.equals(CellType.NUMERIC)) {
				
				double nCellValue = c.getNumericCellValue();
				
				System.out.println(nCellValue);
			}
		}
}
	public static void Cell_Data() throws IOException {
		File f2 = new File("C:\\Users\\User\\eclipse-Maven_NewProject\\Maven_Project\\DataDriven.xlsx");

		FileInputStream fi = new FileInputStream(f2);
		
		Workbook wb1 = new XSSFWorkbook(fi);
		
		Sheet sheetAt = wb1.getSheetAt(0);
		
		int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
		
		for (int i = 0; i < physicalNumberOfRows; i++) {
			
			Row r1 = sheetAt.getRow(i);
			Cell cell = r1.getCell(1);
			
			CellType cellType = cell.getCellType();
			
			if(cellType.equals(CellType.STRING)) {
				
				String stringCellValue = cell.getStringCellValue();
				
				System.out.println(stringCellValue);
			}
			
			else if(cellType.equals(CellType.NUMERIC)) {
				double numericCellValue = cell.getNumericCellValue();
				
				System.out.println(numericCellValue);
			}
	}
		

	}
	
	public static void main(String[] args) throws IOException {
		Particular_Data();
		All_Data();
		Row_Data();
		Cell_Data();
	}
	
	
	
}
