package com.oprataionFile.util;

import org.junit.Before;
import org.junit.Test;

import com.opretaionFile.util.CommonUtil;
import com.opretaionFile.util.ReadFileUtil;

public class TestCommonUtil {
	
	@Before
	public void init() {
		ReadFileUtil instance = ReadFileUtil.getInstance();
		instance.readPropertiesFile();
	}
	
	@Test
	public void testReadExcelRowToFile() {
//		CommonUtil.readExcelRowToFile(columnIdx, filePath, unReadRowSet, sheetIdx);
	}
}
