package javapackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {

		// Specify the location of Excel File
		File src = new File(".\\src\\test\\resources\\datafiles\\Demo.xlsx");

		// Load File
		FileInputStream fis = new FileInputStream(src);

		// Load Workbook
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		// Load WorkSheet
		XSSFSheet sh = wb.getSheet("DemoSheet");

		// Print the name of loaded sheet
		System.out.println(sh.getSheetName());

		// Print UserName from Excel Sheet
		System.out.println(sh.getRow(0).getCell(0).getStringCellValue());

		// Print p2 from Excel Sheet

		System.out.println(sh.getRow(2).getCell(1).getStringCellValue());

		// Print Total Number of Rows
		System.out.println("Total Rows :- " + sh.getPhysicalNumberOfRows());

		// Print Total Number of Columns
		System.out.println("Total Columns:- " + sh.getRow(0).getPhysicalNumberOfCells());

		int rows = (sh.getLastRowNum() - sh.getFirstRowNum()) + 1;
		System.out.println("Total Rows:- " + rows);

		int columns = sh.getRow(0).getLastCellNum();
		System.out.println("Total Columns:- " + columns);

		// Print All Cells of Excel Sheet

		for (int i = 101; i < rows; i++) {

			for (int j = 0; j < columns; j++) {

				System.out.println(sh.getRow(i).getCell(j).getStringCellValue());
			}
		}

	}

}
