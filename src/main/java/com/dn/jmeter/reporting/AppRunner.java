package com.dn.jmeter.reporting;

import java.io.IOException;

public class AppRunner {
	public static void main(String[] args) throws IOException {
		String csvFile = "C:\\Users\\duongkyo\\Downloads\\Report.csv";
		String excelFile = "C:\\Users\\duongkyo\\Downloads\\Report.xlsx";
		GenerateJmeterReport gr = new GenerateJmeterReport();
		gr.copyCsvToExcel(csvFile, excelFile);
		gr.groupRows(excelFile);
	}
}
