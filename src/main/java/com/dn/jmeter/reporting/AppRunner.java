package com.dn.jmeter.reporting;

import java.io.IOException;

public class AppRunner {
	public static void main(String[] args) throws IOException {
		GenerateJmeterReport gr = new GenerateJmeterReport();
		String type = args[0];
		String arg1 = args[1];
		String arg2 = args[2];
		
		if (type.equalsIgnoreCase("CollectLog")) {
			gr.collectLog(arg1, arg2);
		}
		
		if (type.equalsIgnoreCase("Generate")) {
			gr.copyCsvToExcel(arg1, arg2);
			gr.groupRows(arg2);			
		}
	}
}