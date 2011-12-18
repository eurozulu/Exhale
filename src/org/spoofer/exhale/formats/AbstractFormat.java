package org.spoofer.exhale.formats;

import java.io.OutputStream;
import java.io.PrintStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.spoofer.exhale.utils.NamedArgs;

public abstract class AbstractFormat implements ExhaleFormat
{
	protected PrintStream output;
	
	
	public void setArguments(NamedArgs args)
	{
	}

	public void setOutput(OutputStream out)
	{
		this.output = out instanceof PrintStream ? (PrintStream)out : new PrintStream(out);
	}

	
	protected String getCellValue(XSSFCell cell)
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

	
	protected String getCellType(int cellType)
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
