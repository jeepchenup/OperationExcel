package com.oprataionFile.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Before;
import org.junit.Test;

import com.operationFile.model.Config;
import com.opretaionFile.util.ReadFileUtil;

public class TestReadFileUtil {

	private File readFile = null;
	private File updateFile = null;
	private File templateFile = null;
	private ReadFileUtil instance = null;

	@Before
	public void setConfiguration() throws EncryptedDocumentException, InvalidFormatException, IOException {
		instance = ReadFileUtil.getInstance();
		instance.readPropertiesFile();
		Config config = instance.readPropertiesFile();
		updateFile = ReadFileUtil.findFile(config.getResultFilePath());
		readFile = ReadFileUtil.findFile(config.getReadFilePath());
		templateFile = ReadFileUtil.findFile(config.getTemplateFilePath());
	}

	@Test
	public void testFindFile() {
		updateFile.exists();
	}
	
	@Test
	public void testUpdateXLSFile() throws EncryptedDocumentException, InvalidFormatException, IOException {
		instance.updateXLSFile(updateFile, null);
	}
	
	@Test
	public void testGetSheetInfoOfWorkbook() throws EncryptedDocumentException, InvalidFormatException, IOException {
		FileInputStream fis = new FileInputStream(updateFile);
		Workbook workbook = WorkbookFactory.create(fis);
		fis.close();
		workbook.close();
		
		int totalLength = workbook.getNumberOfSheets();
		Map info = instance.getSheetInfoOfWorkbook(workbook);
		
		System.out.println("total sheet number : " + totalLength);
		for(int i = 0; i<totalLength; i++) {
			System.out.println(info.get(i));
		}
	}
	
	@Test
	public void testReadXLSFileAsMap() throws IOException {
		Map<Object, Object> map = instance.readXLSFileAsMap(readFile);
		Map mps = (Map) map.get("协鑫新能源");
		System.out.println(mps);
	}
	
	@Test
	public void testReadPropertiesFile() {
		instance.readPropertiesFile();
	}
	
	@Test
	public void testCopySheetAsTemplate() throws EncryptedDocumentException, InvalidFormatException, IOException {
		FileInputStream fis = new FileInputStream(readFile);
		Workbook workbook = WorkbookFactory.create(fis);
		fis.close();
		Map<Integer, String> dataMap = instance.getSheetInfoOfWorkbook(workbook);
		workbook.close();
		
		instance.copySheetAsTemplate(updateFile, dataMap);
	}
}
