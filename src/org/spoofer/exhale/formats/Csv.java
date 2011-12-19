package org.spoofer.exhale.formats;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.spoofer.exhale.utils.NamedArgs;

public class Csv extends AbstractFormat implements ExhaleFormat
{

	public static final String DEFAULT_DELIMITER = ",";
	public static final String DEFAULT_LINE_DELIMITER = "\n\r";

	public static final String ARG_EX_HEADERS = "noheaders";
	public static final String ARG_DELIMIT = "delimit";
	public static final String ARG_LINE_DELIMIT = "line";

	private String delimit = DEFAULT_DELIMITER;
	private String lineDelimit = DEFAULT_LINE_DELIMITER;

	private boolean includeHeaders = true;
	
	private boolean firstCell;
	private boolean includeRow;

	public void setArguments(NamedArgs args)
	{
		if (args.contains(ARG_DELIMIT))
			delimit = args.getArgument(ARG_DELIMIT);

		if (args.contains(ARG_LINE_DELIMIT))
			lineDelimit = args.getArgument(ARG_LINE_DELIMIT);

	}


	public void openBook(XSSFWorkbook workbook) throws IOException
	{}

	public void openSheet(XSSFSheet sheet) throws IOException
	{
		output.println();
		output.println(sheet.getSheetName());
	}

	public void openRow(XSSFRow row) throws IOException
	{
		includeRow = includeHeaders || !isHeaderRow(row);
		firstCell = true;
	}

	/**
	 * Write a cell to the output.
	 * Note this cell can be null, in instanes where sheets do not expose a cell at a given index.
	 * 
	 */
	public void writeCell(XSSFCell cell) throws IOException
	{
		if (includeRow)
		{
			if (firstCell)
				firstCell = false;
			else
				output.print(delimit);
			
			if (null != cell)
				output.print(getCellValue(cell));
		}
	}

		
	public void closeRow(XSSFRow row) throws IOException
	{
		if (includeRow)
			output.print(lineDelimit);
	}

	public void closeSheet(XSSFSheet sheet) throws IOException
	{
		output.print(lineDelimit);
	}

	public void closeBook(XSSFWorkbook workbook) throws IOException
	{
	}

}
