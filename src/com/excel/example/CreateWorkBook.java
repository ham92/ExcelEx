package com.excel.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorkBook {

	public static void main(String[] args) throws IOException {

		// createXclFile();
		readFromXclFile();

	}

	public static void readFromXclFile() throws IOException {

		XSSFRow row;
		FileInputStream fileInputStream = new FileInputStream(new File("createworkbook.xlsx"));

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet spreadsheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = spreadsheet.iterator();

		while (rowIterator.hasNext()) {
			row = (XSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue() + " \t\t ");
					break;

				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue() + " \t\t ");
					break;
				}
			}

			System.out.println();
		}
		fileInputStream.close();
	}

	public static void createXclFile() throws IOException {

		// Create Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank spreadsheet
		XSSFSheet spreadsheet = workbook.createSheet("Employee Info");
		XSSFRow row = spreadsheet.createRow(1);

		// Create file system using specific name
		FileOutputStream out = null;

		out = new FileOutputStream(new File("createworkbook.xlsx"));

		Map<String, Object[]> empInfo = new TreeMap<String, Object[]>();
		empInfo.put("1", new Object[] { "EMP ID", "EMP NAME", "DESIGNATION" });
		empInfo.put("2", new Object[] { "tp01", "Gopal", "Technical Manager" });
		empInfo.put("3", new Object[] { "tp02", "Manisha", "Proof Reader" });
		empInfo.put("4", new Object[] { "tp03", "Masthan", "Technical Writer" });
		empInfo.put("5", new Object[] { "tp04", "Satish", "Technical Writer" });
		empInfo.put("6", new Object[] { "tp05", "Krishna", "Technical Writer" });

		System.out.println(empInfo.keySet());

		Set<String> keyid = empInfo.keySet();
		int rowid = 0;

		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);

		for (String key : keyid) {
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = empInfo.get(key);
			int cellid = 0;

			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}

		}
		// write operation workbook using file out object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");

	}

}
