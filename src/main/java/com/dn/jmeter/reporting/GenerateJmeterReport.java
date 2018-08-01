package com.dn.jmeter.reporting;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GenerateJmeterReport {

	public void groupRows(String excelFile) throws IOException {
		File inputFile = new File(excelFile);
		FileInputStream inputStream = new FileInputStream(inputFile);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		Boolean skipFirstRow = true;
		Iterator<Row> rowIterator = sheet.iterator();
		int fromRow = 1;
		while (rowIterator.hasNext()) {
			if (skipFirstRow) {
				skipFirstRow = false;
				continue;
			}
			Row row = rowIterator.next();
			String cellValue = row.getCell(0).getStringCellValue();
			if (!Character.isDigit(cellValue.charAt(0))) {
				sheet.groupRow(fromRow, row.getRowNum() - 1);
				sheet.setRowSumsBelow(true);
				sheet.setRowGroupCollapsed(row.getRowNum(), true);
				fromRow = row.getRowNum() + 1;
			}
		}
		inputStream.close();
		File outputFile = new File(excelFile);
		FileOutputStream outputStream = new FileOutputStream(outputFile);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
	}

	public void copyCsvToExcel(String csv, String excel) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		File csvFile = new File(csv);
		BufferedReader bufRdr = new BufferedReader(new FileReader(csvFile));
		int rowNum = 0;
		String line = null;
		while ((line = bufRdr.readLine()) != null) {
			String[] raw = line.split(",");
			XSSFRow row = sheet.createRow(rowNum);
			for (int i = 0; i < raw.length; i++) {
				row.createCell(i).setCellValue(raw[i]);
			}
			rowNum++;
		}
		bufRdr.close();
		FileOutputStream outputStream = new FileOutputStream(excel);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
	}
	
	public static void main(String[] args) throws IOException {
		String csvFile = "C:\\Users\\duongkyo\\Downloads\\Report.csv";
		String excelFile = "C:\\Users\\duongkyo\\Downloads\\Report.xlsx";
		GenerateJmeterReport gr = new GenerateJmeterReport();
		gr.copyCsvToExcel(csvFile, excelFile);
		gr.groupRows(excelFile);
	}
}
