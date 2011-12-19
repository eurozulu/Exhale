package org.spoofer.exhale;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.spoofer.exhale.formats.ExhaleFormat;
import org.spoofer.exhale.formats.FormatException;
import org.spoofer.exhale.formats.FormatFactory;
import org.spoofer.exhale.utils.NamedArgs;
import org.spoofer.exhale.utils.Strings;

public class Exhale
{
	public static final String DEFAULT_FORMAT = "csv";

	public static final String ARG_TO = "to";
	public static final String ARG_SHEET = "sheet";

	public static final String ARG_FILE = null;

	public static final String ARG_FORMATS = "formats";
	private static final String ARG_TIME = "t";

	private static final CharSequence SHEET_NAME_DELIMITER = ":";

	
	private ExhaleFormat format;
	
	private String[] sheets = null;


	public void exhaleStream(InputStream in, OutputStream out) throws IOException, FormatException
	{
		if (null == getFormat())
			throw new FormatException("No format set to convert to.  Use 'setFormat' to specify the required output.");
		
		format.setOutput(out);

		XSSFWorkbook workbook;
		try {
			workbook = new XSSFWorkbook(in);
		} catch (IOException e) {
			throw new IllegalArgumentException("Failed to open file as xslt work book");
		}

		format.openBook(workbook);

		int count = workbook.getNumberOfSheets();
		for (int sheetIndex = 0; sheetIndex < count; sheetIndex++)
		{
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			
			if (isNamedSheet(sheet))
			{
				format.openSheet(sheet);

				int rowCount = sheet.getLastRowNum();
				for (int rowIndex=0; rowIndex < rowCount; rowIndex++)
				{
					XSSFRow row = sheet.getRow(rowIndex);
					format.openRow(row);

					int cellCount = row.getLastCellNum();
					for (int cellIndex=0; cellIndex < cellCount; cellIndex++)
					{
						XSSFCell cell = row.getCell(cellIndex);
						format.writeCell(cell);
					}

					format.closeRow(row);
				}
				
				format.closeSheet(sheet);
			}
		}

		format.closeBook(workbook);
	}


	public String getFormat()
	{ return null != this.format ? format.getClass().getSimpleName() : null;}

	public void setFormat(String format, NamedArgs args) throws FormatException
	{
		if (!Strings.isEmpty(format)) {
			
			this.format = FormatFactory.loadFormat(format);
			if (null != args)
				this.format.setArguments(args);
			
		} else
			throw new IllegalArgumentException("format can not be null or empty");
	}


	public String getSheet()
	{ 
		if (null == this.sheets)
			return null;

		StringBuilder s = new StringBuilder();
		for (String name : this.sheets)
		{
			if (s.length() > 0)
				s.append(SHEET_NAME_DELIMITER);
			s.append(name);
		}

		return s.toString();
	}

	/**
	 * Sets a list of all the sheet names to include in the export.
	 * By default, all sheets are exported.
	 * By naming one or more sheets, only the named sheets are exported.
	 * Sheet names should be serarated using a colon ':', e.g sheet1:sheet2:sheet3 etc.
	 * Note naming the sheets does not alter the order in which the output is written.
	 * Output will be as it appears in the source file.
	 * 
	 * @param sheet A single sheet name or multiple, colon delimited sheets names.
	 * 				Setting to null will revert to displaying all sheets.
	 * 
	 */
	public void setSheets(String sheet)
	{
		if (Strings.isEmpty(sheet))
			this.sheets = null;
		else 
			this.sheets = sheet.contains(SHEET_NAME_DELIMITER) ? sheet.split("\\" + SHEET_NAME_DELIMITER) : new String[]{sheet};

	}
	/**
	 * Checks if the given name is in the list of sheets to export.
	 * If the sheets names have been set, then the given name is checked against those names.
	 * If the names match this returns true.
	 * This also returns true when no sheet names have been set.
	 * If names have been set and the given name is not n those set names, this returns false.
	 * @param sheet
	 * @return
	 */
	public boolean isNamedSheet(XSSFSheet sheet)
	{
		if (null == this.sheets)
			return true;
		else if (null == sheet)
			return false;

		boolean found = false;
		for (String name : this.sheets)
		{
			found = ( (Strings.isNumber(name) && sheet.getWorkbook().getSheetIndex(sheet) == Integer.parseInt(name)) ||
					  sheet.getSheetName().equalsIgnoreCase(name) );
			if (found)
				break;
		}

		return found;
	}

	/**
	 * @param args
	 */
	public static void main(String[] rawargs)
	{
		
		long startTime = System.currentTimeMillis();
		
		NamedArgs args = new NamedArgs(rawargs);
		// TODO: Fix the class loading of the avaliable formats
		if (args.contains(ARG_FORMATS))
		{
			System.out.println(Arrays.toString(FormatFactory.getFormats()) );
			System.exit(0);
		}
		InputStream in;
		String fileName = args.getArgument(ARG_FILE);
		try {

			in = !Strings.isEmpty(fileName) ? new FileInputStream(fileName) : System.in;

		} catch (FileNotFoundException e) {
			throw new IllegalArgumentException("failed to open file " + fileName, e);
		}

		OutputStream out = System.out;

		Exhale ex = new Exhale();

		
		String format = args.getArgument(ARG_TO);
		if (Strings.isEmpty(format))
			format = DEFAULT_FORMAT;
		try {
				ex.setFormat(format, args);
		} catch (FormatException e) {
			throw new IllegalArgumentException("can not format into " + args.getArgument(ARG_TO), e);
		}
		
		if (args.contains(ARG_SHEET))
			ex.setSheets(args.getArgument(ARG_SHEET));
		
		try {
			ex.exhaleStream(in, out);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (FormatException e) {
			e.printStackTrace();
		}

		long endTime = System.currentTimeMillis();
		
		if (args.contains(ARG_TIME))
			System.out.println("\nTook:" + (endTime - startTime) + " miliseconds");
	}


}
