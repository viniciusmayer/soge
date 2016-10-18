package com.eleonorvinicius.soge.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {

	private static final String RS = "RS";
	private static final String UF = "UF";
	private static final String FILES_DIRECTORY = "/home/eleonorvinicius/Projects/soge/files/";

	public static void main(String[] args) {
		File directory = new File(FILES_DIRECTORY);
		String[] list = directory.list();
		for (String file : list) {
			Sheet sheet = getSheet(file);
			Row row = sheet.getRow(1);
			Integer ufColumn = getUFColumn(row);
			Iterator<Row> rowIterator = sheet.iterator();
			int rowCount = 0;
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				if (rowCount < 2) {
					rowCount += 1;
					continue;
				}
				StringBuilder linha = new StringBuilder("");
				Cell cell = row.getCell(ufColumn);
				if (cell != null){
					String cellValue = getCellValue(cell);
					if (cellValue != null && cellValue.equals(RS)){
						Iterator<Cell> cellIterator = row.cellIterator();
						while(cellIterator.hasNext()){
							Cell _cell = cellIterator.next();
							String s = getCellValue(_cell);
							linha.append(s);
							linha.append(" | ");
						}
						System.out.println(linha);
					}
				}
			}
		}
	}
	
	private static Sheet getSheet(String file){
		FileInputStream fileInputStream = null;
		try {
			fileInputStream = new FileInputStream(new File(FILES_DIRECTORY + file));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return null;
		}
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(fileInputStream);
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
			return null;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			return null;
		} catch (IOException e) {
			e.printStackTrace();
			return null;
		}
		return workbook.getSheetAt(0);
	}

	private static String getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}

	private static Integer getUFColumn(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		Integer ufColumn = null;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			String s = null;
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				s = cell.getStringCellValue();
				break;
			}
			if (s != null && s.equals(UF)) {
				ufColumn = cell.getColumnIndex();
				break;
			}
		}
		return ufColumn;
	}
}