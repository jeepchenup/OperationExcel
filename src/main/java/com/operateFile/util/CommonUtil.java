package com.opretaionFile.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.SystemOutLogger;

import com.operationFile.constants.ConfigConstants;
import com.operationFile.constants.UtilConstants;

public class CommonUtil {

	/**
	 * 通过指定的列数，获取Excel数据
	 * @param columnIdx 指定读取列数
	 * @param filePath 读取文件路径
	 * @param unReadRowSet 不读取行数集合
	 * @param sheetIdx 指定读取sheet, 开始为0
	 * @param startRowIdx 开始读取的行数
	 * @param endRowIdx 结束读取的行数
	 */
	public static void readExcelRowToFile(int columnIdx, String filePath, String unReadRowSet, int sheetIdx, int startRowIdx, int endRowIdx) {
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
				int currentRowIdx = row.getRowNum();
				if( !unreadSet.contains(currentRowIdx) && currentRowIdx >= startRowIdx && currentRowIdx <= endRowIdx) {
					getInfo.append(row.getCell(columnIdx).getStringCellValue())
						   	   .append(UtilConstants.NEW_LINE_CHAR);
					if(currentRowIdx == endRowIdx) break;
				}
			}
			read.close();
			workbook.close();
			
			StringBuffer outputFileName = new StringBuffer();
			StringBuffer outputFilePath = new StringBuffer();
			
			outputFileName.append(formatDate(Calendar.getInstance().getTime()))
						  .append(ConfigConstants.DASH)
						  .append(file.getName().split("\\.")[0])
						  .append(ConfigConstants.TXT_TYPE);
			outputFilePath.append(System.getProperty("user.dir"))
						  .append(ConfigConstants.OUTPUT_FILE_PATH)
						  .append(outputFileName);
			
			File outputFile = new File(outputFilePath.toString());
			if(!outputFile.exists())  {
				outputFile.getParentFile().mkdir();
				outputFile.createNewFile();
			}
			out = new FileOutputStream(outputFile);
			out.write(getInfo.toString().getBytes());
			out.flush();
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
	
	public static String formatDate(Date date) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		return sdf.format(date);
	}
	
	/**
	 * 获取不能读取行数的List
	 * @param setStr 不能读取的字符串集合。例如："1,2,3"
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
