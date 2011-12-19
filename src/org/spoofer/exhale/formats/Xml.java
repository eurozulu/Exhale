package org.spoofer.exhale.formats;

import java.io.IOException;
import java.io.OutputStream;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.spoofer.exhale.Exhale;
import org.spoofer.exhale.utils.NamedArgs;
import org.spoofer.exhale.utils.Strings;

import com.sun.xml.internal.txw2.output.IndentingXMLStreamWriter;
import com.sun.xml.internal.ws.api.streaming.XMLStreamWriterFactory;

public class Xml extends AbstractFormat implements ExhaleFormat
{

	public static final String OUTPUT_NAMESPACE = Exhale.class.getPackage().getName() + "." + Xml.class.getSimpleName().toLowerCase();

	public static final String ENCODING = "UTF8";
	public static final String TAG_BOOK = "workbook";
	public static final String TAG_SHEET = "sheet";
	public static final String TAG_ROW = "row";
	public static final String TAG_CELL = "c";

	public static final String ATTR_SHEET_NAME = "name";
	public static final String ATTR_ROW_INDEX = "rowIndex";
	public static final String ATTR_COL_INDEX = "colIndex";
	public static final String ATTR_TYPE = "type";

	public static final String ARG_INCLUDE_BLANKS = "blanks";
	public static final String ARG_INDENT = "indent";


	
	private XMLStreamWriter xmlOutput;

	private boolean includeBlankCells = false;
	private boolean indentOutput = false;

	public void openBook(XSSFWorkbook workbook) throws IOException
	{
		try {
			xmlOutput.setDefaultNamespace(OUTPUT_NAMESPACE);
			xmlOutput.writeStartDocument();

			xmlOutput.writeStartElement(OUTPUT_NAMESPACE, TAG_BOOK);

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}


	public void openSheet(XSSFSheet sheet) throws IOException
	{
		try {

			xmlOutput.writeStartElement(TAG_SHEET);
			xmlOutput.writeAttribute(ATTR_SHEET_NAME, sheet.getSheetName());

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}

	}

	public void openRow(XSSFRow row) throws IOException
	{
		try {
			xmlOutput.writeStartElement(TAG_ROW);
			xmlOutput.writeAttribute(ATTR_ROW_INDEX, Integer.toString(row.getRowNum()));

		}catch(XMLStreamException e) {
			throw new IOException(e);
		}
	}

	public void writeCell(XSSFCell cell) throws IOException
	{
		try {
			String value = getCellValue(cell);
			
			if (!Strings.isEmpty(value)) {
				
				xmlOutput.writeStartElement(TAG_CELL);
				xmlOutput.writeAttribute(ATTR_TYPE, getCellType(cell.getCellType()));
				xmlOutput.writeAttribute(ATTR_COL_INDEX, Integer.toString(cell.getColumnIndex()));
				xmlOutput.writeAttribute(ATTR_ROW_INDEX, Integer.toString(cell.getRowIndex()));

				xmlOutput.writeCharacters(getCellValue(cell));
			
				xmlOutput.writeEndElement();
			} else if (includeBlankCells)
				xmlOutput.writeEmptyElement(TAG_CELL);
			
		}catch(XMLStreamException e) {
			throw new IOException(e);
		}
	}


	public void closeRow(XSSFRow row) throws IOException
	{
		try {
			xmlOutput.writeEndElement();

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	public void closeSheet(XSSFSheet sheet) throws IOException
	{
		try {
			xmlOutput.writeEndElement();

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	public void closeBook(XSSFWorkbook workbook) throws IOException
	{
		try {
			xmlOutput.writeEndElement();

			xmlOutput.writeEndDocument();

			output.println();

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}


	@Override
	public void setOutput(OutputStream out)
	{
		super.setOutput(out);

		this.xmlOutput = XMLStreamWriterFactory.create(out);
		
		if (this.indentOutput)
		{
			this.xmlOutput = new IndentingXMLStreamWriter(xmlOutput);
			//((IndentingXMLStreamWriter)xmlOutput).setIndentStep("\t");
		}
	}
	

	@Override
	public void setArguments(NamedArgs args)
	{
		super.setArguments(args);

		this.includeBlankCells = args.contains(ARG_INCLUDE_BLANKS);
		
		if (args.contains(ARG_INDENT))
		{
			this.indentOutput = true;
			if (null != this.xmlOutput)  // If output already set, reset so it picks up indent wrapper.
				setOutput(output);
		}
	}


}
