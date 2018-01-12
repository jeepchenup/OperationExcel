package com.opretaionFile.util;

import org.junit.Before;
import org.junit.Test;

import com.operationFile.model.Config;
import com.opretaionFile.util.CommonUtil;
import com.opretaionFile.util.ReadFileUtil;

public class TestCommonUtil {
	
	private ReadFileUtil instance;
	private Config config;
	
	@Before
	public void init() {
		instance = ReadFileUtil.getInstance();
		config = instance.readPropertiesFile();
	}
	
	@Test
	public void testReadExcelRowToFile() {
		CommonUtil.readExcelRowToFile(config.getReadColumnIndex(), 
								      config.getTemplateFilePath(),
								      config.getUnreadRowSet(), 
								      config.getReadSheetIndex(),
								      config.getStartReadRowIndex(),
								      config.getEndReadRowIndex());
	}
}
