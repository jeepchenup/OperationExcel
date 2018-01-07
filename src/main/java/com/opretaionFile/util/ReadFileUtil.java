package com.opretaionFile.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.operationFile.constants.ConfigConstants;
import com.operationFile.constants.UtilConstants;
import com.operationFile.model.Config;

/**
 * 
 * @author steven.chen
 * @version 0.2
 *
 */
public class ReadFileUtil {

	private Config innerConfig;
	
	private static ReadFileUtil readFile;
	private ReadFileUtil(){};
	
	public static ReadFileUtil getInstance() {
		if(readFile == null) 
			readFile = new ReadFileUtil();
		return readFile;
	}
	
	public static File findFile(String path) {
		File file = new File(path);
		if(file.exists()) {
			System.out.println(">>>Find your file : " + file.getName());
			return file;
		} else {
			System.out.println(">>>Can not find the file : " +  file.getName() + "\n>>>Please reconfirm the file path!");
			return null;
		}
	}
	
	public static boolean isEmptyData(String data) {
		boolean isEmpty = false;
		if("".equals(data) || "--".equals(data) || (null == data))
			isEmpty = true;
		return isEmpty;
	}
	
	public void updateXLSFile(File file, Map<String, Map<Integer, String>> maps) {
		
		FileInputStream read = null;
		FileOutputStream out = null;
		Workbook workbook = null;
		
		try {
			read = new FileInputStream(file);
			
			workbook = WorkbookFactory.create(read);
			read.close();
			
			if(maps != null) {
				//Loop through the maps
				Iterator<String> mapsIterator = maps.keySet().iterator();
				while(mapsIterator.hasNext()) {
					Map<Integer, String> mapInfo = maps.get(mapsIterator.next());
					
					//TODO insert maps which in the maps to result file
					
				}
			}
			Sheet sheet = workbook.createSheet("sample sheet2");
			
			Row row = sheet.createRow(1);
			Cell cell1 = row.createCell(2);
			Cell cell2 = row.createCell(3);
			cell1.setCellValue("v1");
			cell2.setCellValue("v2");
			
			out = new FileOutputStream(file);
			workbook.write(out);
			workbook.close();
			out.close();
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
	
	public Map<Integer, String> getSheetInfoOfWorkbook(Workbook workbook) {
		
		Map<Integer, String> sheetMap = new HashMap<Integer, String>();
		
		for(int i =0, totalSheets =  workbook.getNumberOfSheets(); i < totalSheets; i++) {
			sheetMap.put(i,workbook.getSheetName(i));
		}
		
		return sheetMap;
	}
	
	public Map<Object, Object> readXLSFileAsMap(File file) throws IOException {
		
		Workbook workbook;
		HSSFSheet sheet;
		HSSFRow row; 
		HSSFCell keyCell;
		HSSFCell valueCell;
		Iterator rows;
		int totalSheetNum = 0;
		
		FileInputStream fis = new FileInputStream(file);
		workbook = new HSSFWorkbook(fis);
		Map sheetsNameInfoMap = getSheetInfoOfWorkbook(workbook);
		Map<Object, Object> sheetsDataInfoMap = new HashMap<Object, Object>();
		
		totalSheetNum = workbook.getNumberOfSheets();
		
		//start reading file
		for(int sheetIdx = 0; sheetIdx < totalSheetNum; sheetIdx ++) {
//		for(int sheetIdx = 0; sheetIdx < 1; sheetIdx ++) {
			System.out.println("Start operating sheet: " + (sheetIdx+1) + " - " + sheetsNameInfoMap.get(sheetIdx));
			
			//save the data map
			Map<Object, Object> sheetDataMap  = new HashMap<Object, Object>();
			
			sheet = (HSSFSheet) workbook.getSheetAt(sheetIdx);
			rows = sheet.rowIterator();
			
			String key;
			Double value = 0.0;
			
			while(rows.hasNext()) {
				boolean isValueCellEmpty = false;
				
				row = (HSSFRow) rows.next();
				
				keyCell = row.getCell(innerConfig.getKeyColumnIndex());
				valueCell = row.getCell(innerConfig.getValueColumnIndex());
				
				key = keyCell.getStringCellValue();
				
				if(null != valueCell) {
					switch(valueCell.getCellType()) {
						case HSSFCell.CELL_TYPE_STRING:
							if(isEmptyData(valueCell.getStringCellValue()))
								isValueCellEmpty = true;
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:
							value = valueCell.getNumericCellValue();
							break;
					}
				}
				
				if(isValueCellEmpty || null == valueCell) {
					sheetDataMap.put(key, UtilConstants.EMPTY_VALUE);
//					System.out.println(key + " " + EMPTY_VALUE);
				} else {
					sheetDataMap.put(key, value);
//					System.out.println(key + " " + value);
				}
			}
			sheetsDataInfoMap.put(sheetsNameInfoMap.get(sheetIdx), sheetDataMap);
			System.out.println("Getted sheet : " + sheetsNameInfoMap.get(sheetIdx));
		}
		
		workbook.close();
		fis.close();
		
		return sheetsDataInfoMap;
	}
	
	public Config readPropertiesFile() {
		String propertiesFilePath = System.getProperty("user.dir") + ConfigConstants.CONFIGU_FILE_PATH;
		Config configModel = new Config();
		File config = findFile(propertiesFilePath);
		
		try {
			if(config != null) {
				Properties prop = new Properties();
				FileInputStream read = new FileInputStream(config);
				prop.load(read);
				Iterator<String> iterator = prop.stringPropertyNames().iterator();
				while(iterator.hasNext()) {
					String key = iterator.next();
					if(ConfigConstants.READ_FILE_PATH.equals(key))
						configModel.setReadFilePath(prop.getProperty(key));
					else if(ConfigConstants.RESULT_FILE_PATH.equals(key))
						configModel.setResultFilePath(prop.getProperty(key));
					else if(ConfigConstants.TEMPLATE_FILE_PATH.equals(key))
						configModel.setTemplateFilePath(prop.getProperty(key));
					else if(ConfigConstants.KEY_COLUMN_INDEX.equals(key))
						configModel.setKeyColumnIndex(Integer.parseInt(prop.getProperty(key)));
					else if(ConfigConstants.VALUE_COLUMN_INDEX.equals(key))
						configModel.setValueColumnIndex(Integer.parseInt(prop.getProperty(key)));
					else if(ConfigConstants.UNREAD_ROW_SET.equals(key))
						configModel.setUnreadRowSet(prop.getProperty(key));
					else if(ConfigConstants.READ_COLUMN_INDEX.equals(key))
						configModel.setReadColumnIndex(Integer.parseInt(prop.getProperty(key)));
					else if(ConfigConstants.READ_SHEET_INDEX.equals(key))
						configModel.setReadSheetIndex(Integer.parseInt(prop.getProperty(key)));
					else if(ConfigConstants.START_READ_ROW_INDEX.equals(key))
						configModel.setStartReadRowIndex(Integer.parseInt(prop.getProperty(key)));
					else if(ConfigConstants.END_READ_ROW_INDEX.equals(key))
						configModel.setEndReadRowIndex(Integer.parseInt(prop.getProperty(key)));
				}
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		this.innerConfig = configModel;
		return configModel;
	}
	
	public void copySheetAsTemplate(File mergeFile, Map<Integer, String> dataMap) {
		FileInputStream read = null;
		FileOutputStream out = null;
		HSSFWorkbook workbook = null;
		
		Entry<Integer, String> sheetEntry = null;
		String sheetName = null;
		
		try {
			read = new FileInputStream(mergeFile);
			workbook = (HSSFWorkbook) WorkbookFactory.create(read);
			read.close();
			
			//create sheet
			Iterator<Entry<Integer, String>> iterator = dataMap.entrySet().iterator();
			
			while(iterator.hasNext()) {
				sheetEntry = iterator.next();
				workbook.cloneSheet(ConfigConstants.TEMPALTE_SHEET_INDEX);
				sheetName = UtilConstants.PRE_SHEET_NAME + sheetEntry.getValue();
				workbook.setSheetName(sheetEntry.getKey() + 1, sheetName);
				
				System.out.println(">>>copy finish :  sheet index: " + sheetEntry.getKey()  + " - [" + sheetName + "]");
			}
			out = new FileOutputStream(mergeFile);
			workbook.write(out);
			workbook.close();
			out.close();
			
			System.out.println("\n>>>finish copy!");
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
}
