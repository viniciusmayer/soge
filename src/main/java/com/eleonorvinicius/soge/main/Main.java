package com.eleonorvinicius.soge.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
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

	private static final String FILES_DIRECTORY = "/home/eleonorvinicius/Projects/soge/files/";

	public static void main(String[] args) {
		File directory = new File(FILES_DIRECTORY);
		String[] list = directory.list();
		StringBuilder arquivo = new StringBuilder("");
		for (String file : list) {
			Sheet sheet = getSheet(file);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				StringBuilder linha = new StringBuilder("");
				Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()){
					Cell cell = cellIterator.next();
					String s = getCellValue(cell);
					if (s != null){
						s = s.replaceAll("[\n\r]", "");
					}
					linha.append(s);
					linha.append(";");
				}
				arquivo.append(linha + "\n");
			}
		}
		try {
			FileWriter fileWriter = new FileWriter("/home/eleonorvinicius/Projects/soge/files/file.csv");
			fileWriter.write(arquivo.toString());
			fileWriter.close();
		} catch (IOException e) {
			e.printStackTrace();
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
}