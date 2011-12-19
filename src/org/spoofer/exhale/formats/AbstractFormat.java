package org.spoofer.exhale.formats;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.spoofer.exhale.utils.NamedArgs;
import org.spoofer.exhale.utils.Strings;

public abstract class AbstractFormat implements ExhaleFormat
{
	
	/**
	 * A positive number from zero -> the number of rows in a sheet.
	 * Specifies how many rows, from the top row, are considured to be header rows.
	 * Header rows are rows containing field labels, rather than field data.
	 * The headers can be applied to the corresponding data in the non header rows.
	 * 
	 * Usually, this is left at the default of one, which includes just the first row as a header.
	 * It can be adjusted to include further rows or to 0 (or negative) number to treat all rows equally
	 * and none as headers.
	 */
	private static final String ARG_HEADER_ROWS = "headerrows";

	private static final int DEFAULT_TITLE_ROW = 1;
	
	protected PrintStream output;
	private List<String> headerNames;
	
	private int headerRow = DEFAULT_TITLE_ROW;
	
	@Override
	public void setArguments(NamedArgs args)
	{
		if (args.contains(ARG_HEADER_ROWS))
		{
			try {
				headerRow = Integer.parseInt(args.getArgument(ARG_HEADER_ROWS));
			} catch (NumberFormatException e) {
				System.err.println("Argument " + ARG_HEADER_ROWS + " can not parse the value " + args.getArgument(ARG_HEADER_ROWS) + " into a number. Reverting to default value of " + DEFAULT_TITLE_ROW);
				e.printStackTrace();
				headerRow = DEFAULT_TITLE_ROW;
			}
		}
	}

	@Override
	public void setOutput(OutputStream out)
	{
		this.output = out instanceof PrintStream ? (PrintStream)out : new PrintStream(out);
	}
	public OutputStream getOutput()
	{
		return this.output;
	}
	
	@Override
	public void setHeaderRow(int rowIndex) {
		this.headerRow = rowIndex;
	}
	
	@Override
	public int getHeaderRows() {
		return this.headerRow;
	}
	
	@Override
	public void openSheet(XSSFSheet sheet) throws IOException
	{
		if (null != this.headerNames)
			this.headerNames.clear();
		else
			this.headerNames = new ArrayList<String>();
	}

	@Override
	public void closeSheet(XSSFSheet sheet) throws IOException
	{
		if (null != this.headerNames)
		{
			this.headerNames.clear();
			this.headerNames = null;
		}
	}


	
	
	@Override
	public void openRow(XSSFRow row) throws IOException
	{
		if (isHeaderRow(row))  // If header row, capture the cell values as header names.
		{
			int count = row.getLastCellNum();
			for (int index=0; index < count; index++)
			{
				XSSFCell cell = row.getCell(index);
				String value = getCellValue(cell);
				if (!Strings.isEmpty(value))
				{
					if (headerNames.size() < index)  // Overwrite exiting values (from previous rows)
						headerNames.add(value);
					else
						headerNames.add(index, value);
				}
			}
		}
		
	}

	
	
	
	protected boolean isHeaderRow(XSSFRow row)
	{
		if (null == row)
			return false;
		
		return row.getRowNum() < getHeaderRows();
			
	}
	protected String getHeaderName(int index)
	{
		return (null != headerNames) && index < headerNames.size() ? headerNames.get(index) : null;
	}
	
	protected static String getCellValue(XSSFCell cell)
	{
		if (null == cell)
			return null;
		
		String value;

		switch (cell.getCellType())
		{
			case XSSFCell.CELL_TYPE_BOOLEAN : {
				value = cell.getBooleanCellValue() ? "true" : "false";
				break;
			}
			case XSSFCell.CELL_TYPE_NUMERIC : {
				value = Double.toString(cell.getNumericCellValue());
				break;
			}
			case XSSFCell.CELL_TYPE_STRING : {
				value = cell.getStringCellValue();
				break;
			}
			case XSSFCell.CELL_TYPE_FORMULA : {
				value = cell.getCellFormula();
				break;
			}
			default : {
				value = cell.getRawValue();
			}
		}

		return value;
	}

	
	protected static String getCellType(int cellType)
	{
		String type;
		switch(cellType)
		{
		case XSSFCell.CELL_TYPE_BOOLEAN : {
			type = "boolean";
			break;
		}
		case XSSFCell.CELL_TYPE_NUMERIC : {
			type = "numeric";
			break;
		}

		case XSSFCell.CELL_TYPE_STRING : {
			type = "string";
			break;
		}
		case XSSFCell.CELL_TYPE_FORMULA : {
			type = "formula";
			break;
		}

		case XSSFCell.CELL_TYPE_BLANK : {
			type = "blank";
			break;
		}

		case XSSFCell.CELL_TYPE_ERROR : {
			type = "error!";
			break;
		}
		default :
			type = "unknown";

		}
		
		return type;
	}


}
