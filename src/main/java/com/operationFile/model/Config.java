package com.operationFile.model;

public class Config {

	private String readFilePath;
	
	private String resultFilePath;
	
	private String templateFilePath;
	
	private String unreadRowSet;
	
	private int keyColumnIndex;
	
	private int valueColumnIndex;
	
	private int readColumnIndex;
	
	private int readSheetIndex;

	public String getReadFilePath() {
		return readFilePath;
	}

	public String getResultFilePath() {
		return resultFilePath;
	}

	public String getTemplateFilePath() {
		return templateFilePath;
	}

	public String getUnreadRowSet() {
		return unreadRowSet;
	}

	public int getKeyColumnIndex() {
		return keyColumnIndex;
	}

	public int getValueColumnIndex() {
		return valueColumnIndex;
	}

	public int getReadColumnIndex() {
		return readColumnIndex;
	}

	public int getReadSheetIndex() {
		return readSheetIndex;
	}

	public void setReadFilePath(String readFilePath) {
		this.readFilePath = readFilePath;
	}

	public void setResultFilePath(String resultFilePath) {
		this.resultFilePath = resultFilePath;
	}

	public void setTemplateFilePath(String templateFilePath) {
		this.templateFilePath = templateFilePath;
	}

	public void setUnreadRowSet(String unreadRowSet) {
		this.unreadRowSet = unreadRowSet;
	}

	public void setKeyColumnIndex(int keyColumnIndex) {
		this.keyColumnIndex = keyColumnIndex;
	}

	public void setValueColumnIndex(int valueColumnIndex) {
		this.valueColumnIndex = valueColumnIndex;
	}

	public void setReadColumnIndex(int readColumnIndex) {
		this.readColumnIndex = readColumnIndex;
	}

	public void setReadSheetIndex(int readSheetIndex) {
		this.readSheetIndex = readSheetIndex;
	}
	
}
