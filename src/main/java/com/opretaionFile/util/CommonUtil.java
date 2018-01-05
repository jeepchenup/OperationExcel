package com.opretaionFile.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CommonUtil {

	/**
	 * 通过指定的列数，获取Excel数据
	 * @param columnIdx 指定读取列数
	 * @param filePath 读取文件路径
	 * @param unReadRowSet 不读取行数集合
	 * @param sheetIdx 指定读取sheet, 开始为0
	 */
	public static void readExcelRowToFile(int columnIdx, String filePath, String unReadRowSet, int sheetIdx) {
		File file = ReadFileUtil.findFile(filePath);
		FileInputStream read = null;
		FileOutputStream out = null;
		List<Integer> unreadSet = null;
		
		try {
			read = new FileInputStream(file);
			
			Workbook workbook = WorkbookFactory.create(read);
			Sheet sheet = workbook.getSheetAt(sheetIdx);
			unreadSet = parseList(unReadRowSet);
			Iterator<Row> rowsIterator = sheet.rowIterator();
			Row row;
			StringBuffer getInfo = new StringBuffer();
			
			while(rowsIterator.hasNext()) {
				row = rowsIterator.next();
				row.getRowNum();
				
				Cell cell;
				if( !unreadSet.contains(row.getRowNum()) ) {
					cell = row.getCell(columnIdx);
					getInfo.append(cell.getStringCellValue());
				}
			}
			
			read.close();
			workbook.close();
			
			
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
	
	/**
	 * 获取不能读取行数的List
	 * @param setStr 不能读取字符串
	 * @return
	 */
	private static List<Integer> parseList(String setStr) {
		ArrayList<Integer> list = new ArrayList<Integer>();
		String[] array = setStr.split(",");
		for(String unit : array) {
			list.add(Integer.parseInt(unit));
		}
		return list;
	}
}
