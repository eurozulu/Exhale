package org.spoofer.exhale.formats;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.spoofer.exhale.utils.NamedArgs;

public interface ExhaleFormat {

	public void openBook(XSSFWorkbook workbook) throws IOException;

	public void openSheet(XSSFSheet sheet) throws IOException;
	public void openRow(XSSFRow row) throws IOException;

	public void writeCell(XSSFCell cell) throws IOException;

	public void closeRow(XSSFRow row) throws IOException;
	public void closeSheet(XSSFSheet sheet) throws IOException;
	public void closeBook(XSSFWorkbook workbook) throws IOException;

	/**
	 * Sets the command line arguments. All arguments are passed to the
	 * Formatter.
	 * 
	 * @param args
	 */
	public void setArguments(NamedArgs args);

	/**
	 * Gets a display string to present when user wants help screen.
	 * This is combined with all other formats to rpesent the arguments available.
	 * @return
	 */
	public void getArgumentHelp(PrintStream out);
	
	/**
	 * Sets the output stream to receive the export.
	 * 
	 * @param out
	 */
	public void setOutput(OutputStream out);
	public OutputStream getOutput();

	/**
	 * Sets the number of rows to include as Header fields. By default, this is
	 * one, indicating the first row in each sheet contains header fields.
	 * Setting this to a number greater than 1 will include each consecutive row
	 * below the first, upto the count of rows set. e.g. setting this to 3 will
	 * include the first three rows as headers. When more than one of those rows
	 * contains a value, The last row, the one with the highest index, is used.
	 * 
	 * @param rowIndex
	 */
	public void setHeaderRow(int rowIndex);

	/**
	 * Gets the number of rows to use as header fields.
	 * 
	 * @return
	 */
	public int getHeaderRows();
}
