package com.dn.jmeter.reporting;

import java.io.IOException;

public class AppRunner {
	public static void main(String[] args) throws IOException {
		String csvFile = args[0];
		String excelFile = args[1];
		GenerateJmeterReport gr = new GenerateJmeterReport();
		gr.copyCsvToExcel(csvFile, excelFile);
		gr.groupRows(excelFile);
	}
}
