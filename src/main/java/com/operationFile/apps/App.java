package com.operationFile.apps;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.opretaionFile.util.ReadFileUtil;

public class App {

	private ReadFileUtil instance;
	
	private void init() {
		instance = ReadFileUtil.getInstance();
		instance.readPropertiesFile();
	}
	
	public void run() {
		Map<Integer,String> sheetInfosMap;
		FileInputStream read;
		FileOutputStream out;
		File dataFile, mergeFile;
		Workbook dataWorkbook;
		Workbook mergeWorkbook;
		init();
		
		try {
			dataFile = ReadFileUtil.findFile(instance.read_file_path);
			mergeFile = ReadFileUtil.findFile(instance.result_file_path);
			
			read = new FileInputStream(dataFile);
			dataWorkbook = WorkbookFactory.create(read);
			sheetInfosMap = instance.getSheetInfoOfWorkbook(dataWorkbook);
			
			read.close();
			dataWorkbook.close();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	public static void main(String[] args) {
		App app = new App();
		app.run();
	}
}
