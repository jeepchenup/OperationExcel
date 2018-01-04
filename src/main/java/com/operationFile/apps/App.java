package com.operationFile.apps;

import com.opretaionFile.util.ReadFileUtil;

public class App {

	private ReadFileUtil readFileUtil = ReadFileUtil.getInstance();
	
	public void operationExcel() {
		readFileUtil.readPropertiesFile();
		
	}
	
	public static void main(String[] args) {
		
	}
}
