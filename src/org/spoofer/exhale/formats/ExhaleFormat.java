package org.spoofer.exhale.formats;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.spoofer.exhale.utils.NamedArgs;

public interface ExhaleFormat
{

	public void openBook(XSSFWorkbook workbook) throws IOException;
	
	public void openSheet(XSSFSheet sheet) throws IOException;
	public void openRow(XSSFRow row) throws IOException;
	
	public void writeCell(XSSFCell cell) throws IOException;
	
	public void closeRow(XSSFRow row) throws IOException;
	public void closeSheet(XSSFSheet sheet) throws IOException;
	public void closeBook(XSSFWorkbook workbook) throws IOException;
	
	public void setArguments(NamedArgs args);
	public void setOutput(OutputStream out);
}
