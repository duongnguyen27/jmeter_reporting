package com.dn.jmeter.reporting;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GenerateJmeterReport {
	
	public void collectLog(String inputFolder, String outputLog) throws IOException {
		List<String> fullLog = new ArrayList<String>();
		File folder = new File(inputFolder);
		int fileSequence = 0 ;
		for (File file : folder.listFiles()) {
			if (file.getName().endsWith(".jtl")) {
				BufferedReader bufRdr = new BufferedReader(new FileReader(file));
				Boolean ignoreLine = true;
				String line = null;
				while ((line = bufRdr.readLine()) != null) {
					if (fileSequence > 0 && ignoreLine) {
						ignoreLine = false;
						continue;
					}
					fullLog.add(line);
				}
				fileSequence++;
				bufRdr.close();
			}
		}
		
		FileWriter fileWriter = null;
		try {
			fileWriter = new FileWriter(outputLog);
			for (int i = 0; i < fullLog.size(); i++) {
				fileWriter.append(fullLog.get(i));
				fileWriter.append("\n");
			}
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally {
			try {
				fileWriter.flush();
				fileWriter.close();
			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public void copyCsvToExcel(String csv, String excel) throws IOException {
	
		// Read csv file
		File csvFile = new File(csv);
		BufferedReader bufRdr = new BufferedReader(new FileReader(csvFile));
		List<String[]> lines = new ArrayList<String[]>();
		String line = null;
		while ((line = bufRdr.readLine()) != null) {
			String[] raw = line.split(",");
			lines.add(raw);
		}
		bufRdr.close();
	
		// Save the header and footer lines
		List<String[]> headerFooter = new ArrayList<String[]>();
		headerFooter.add(lines.get(0));
		headerFooter.add(lines.get(lines.size() - 1));
	
		// Remove lines that don't starts with index
		for (int i = 0; i < lines.size(); i++) {
			if (!Character.isDigit(lines.get(i)[0].charAt(0))) {
				lines.remove(lines.get(i));
				i = i - 1;
			}
		}
	
		// Sort by index ascending
		Collections.sort(lines, new Comparator<String[]>() {
			public int compare(String[] one, String[] other) {
				int index1 = Integer.parseInt(one[0].substring(0, one[0].indexOf("_")));
				int index2 = Integer.parseInt(other[0].substring(0, other[0].indexOf("_")));
				return ((Integer) index1).compareTo(index2);
			}
		});
	
		// Add back header footer
		lines.add(0, headerFooter.get(0));
		lines.add(headerFooter.get(1));
	
		// Write lines to excel
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Report");
		for (int i = 0; i < lines.size(); i++) {
			XSSFRow row = sheet.createRow(i);
			for (int j = 0; j < lines.get(i).length; j++) {
				row.createCell(j).setCellValue(lines.get(i)[j]);
			}
		}
	
		FileOutputStream outputStream = new FileOutputStream(excel);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
	}

	public void groupRows(String excelFile) throws IOException {
		File inputFile = new File(excelFile);
		FileInputStream inputStream = new FileInputStream(inputFile);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();
		int fromRow = 2; // Index of row to group from
		int skipRows = 2; // Skip the header and first expand-collapse row
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			String cellValue = row.getCell(0).getStringCellValue();
			String regex = "([0-9]+)_";
			Pattern pattern = Pattern.compile(regex);
			Matcher matcher = pattern.matcher(cellValue);
			Boolean found = matcher.find();

			// Group rows
			if ((found && Integer.parseInt(matcher.group(1)) % 1000 == 0) || !found) {
				if (row.getRowNum() > skipRows) {
					sheet.groupRow(fromRow, row.getRowNum() - 1);
					sheet.setRowSumsBelow(false);
					fromRow = row.getRowNum() + 1;
				}
			}

			// Remove index from label
			if (found) {
				row.getCell(0).setCellValue(cellValue.substring(cellValue.indexOf("_") + 1));
			}
		}

		inputStream.close();
		File outputFile = new File(excelFile);
		FileOutputStream outputStream = new FileOutputStream(outputFile);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
	}
}
