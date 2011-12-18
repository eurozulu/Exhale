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
	
	private static final String ARG_TO = "to";

	private static final String ARG_FILE = null;

	private static final String ARG_FORMATS = "formats";
	
	private String format = DEFAULT_FORMAT;
	

	public void exhaleStream(InputStream in, OutputStream out) throws IOException, FormatException
	{
		ExhaleFormat formatter = FormatFactory.loadFormat(this.getFormat());
		formatter.setOutput(out);
		
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		
		formatter.openBook(workbook);
		
		int count = workbook.getNumberOfSheets();
		for (int sheetIndex = 0; sheetIndex < count; sheetIndex++)
		{
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
			formatter.openSheet(sheet);
			
			int rowCount = sheet.getLastRowNum();
			for (int rowIndex=0; rowIndex < rowCount; rowIndex++)
			{
				XSSFRow row = sheet.getRow(rowIndex);
				formatter.openRow(row);
				
				int cellCount = row.getLastCellNum();
				for (int cellIndex=0; cellIndex < cellCount; cellIndex++)
				{
					XSSFCell cell = row.getCell(cellIndex);
					formatter.writeCell(cell);
						
				}
				
				formatter.closeRow(row);
			}
			
			formatter.closeSheet(sheet);
		}
		
		formatter.closeBook(workbook);
	}
	
	
	public String getFormat()
	{ return this.format;}
	
	public void setFormat(String format)
	{
		if (!Strings.isEmpty(format))
			this.format = format;
		else
			throw new IllegalArgumentException("format can not be null or empty");
	}
	
	
	/**
	 * @param args
	 */
	public static void main(String[] rawargs)
	{
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
		
		if (args.contains(ARG_TO))
			ex.setFormat(args.getArgument(ARG_TO));
		
		try {
			ex.exhaleStream(in, out);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (FormatException e) {
			e.printStackTrace();
		}

	}

}
