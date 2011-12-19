package org.spoofer.exhale.formats;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;

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

public class Xml extends AbstractFormat implements ExhaleFormat {

	public static final String OUTPUT_NAMESPACE = Exhale.class.getPackage()
			.getName()
			+ "." + Xml.class.getSimpleName().toLowerCase();

	public static final String ENCODING = "UTF8";
	public static final String TAG_BOOK = "workbook";
	public static final String TAG_SHEET = "sheet";
	private static final String TAG_HEADER = "header";
	public static final String TAG_ROW = "row";
	public static final String TAG_CELL = "c";

	public static final String ATTR_SHEET_NAME = "name";
	public static final String ATTR_ROW_INDEX = "rowIndex";
	public static final String ATTR_COL_INDEX = "colIndex";
	public static final String ATTR_HEADER_NAME = "headername";
	public static final String ATTR_TYPE = "type";

	public static final String ARG_INCLUDE_BLANKS = "blanks";
	public static final String ARG_INDENT = "indent";
	public static final String ARG_INCLUDE_HEADNAMES = "headernames";
	public static final String ARG_XINDEX = "noindex";
	public static final String ARG_XTYPE = "notype";

	protected XMLStreamWriter xmlOutput;

	private boolean includeBlankCells = false;
	private boolean indentOutput = false;

	private boolean includeHeaderNames = true;
	private boolean excludeIndex = true;
	private boolean excludeType = true;

	public void openBook(XSSFWorkbook workbook) throws IOException {
		try {
			xmlOutput.setDefaultNamespace(OUTPUT_NAMESPACE);
			xmlOutput.writeStartDocument();

			xmlOutput.writeStartElement(OUTPUT_NAMESPACE, TAG_BOOK);

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	public void openSheet(XSSFSheet sheet) throws IOException {
		super.openSheet(sheet);

		try {

			xmlOutput.writeStartElement(TAG_SHEET);
			xmlOutput.writeAttribute(ATTR_SHEET_NAME, sheet.getSheetName());

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}

	}

	@Override
	public void openRow(XSSFRow row) throws IOException {

		try {
			String tag = isHeaderRow(row) ? TAG_HEADER : TAG_ROW;

			xmlOutput.writeStartElement(tag);
			
			if (!this.excludeIndex)
				xmlOutput.writeAttribute(ATTR_ROW_INDEX, Integer.toString(row.getRowNum()));

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	@Override
	public void writeCell(XSSFCell cell) throws IOException {
		super.writeCell(cell);

		try {
			String value = getCellValue(cell);

			if (!Strings.isEmpty(value)) {

				xmlOutput.writeStartElement(TAG_CELL);
				if (value.startsWith("407"))
					System.out.println();
				
				if (!excludeType)
					xmlOutput.writeAttribute(ATTR_TYPE, getCellType(cell
							.getCellType()));

				if (!excludeIndex) {
					xmlOutput.writeAttribute(ATTR_COL_INDEX, Integer
							.toString(cell.getColumnIndex()));
					xmlOutput.writeAttribute(ATTR_ROW_INDEX, Integer
							.toString(cell.getRowIndex()));
				}

				if (includeHeaderNames && !isHeaderRow(cell.getRow())) {
					String name = getHeaderName(cell.getColumnIndex());
					if (Strings.isEmpty(name))
						name = "";

					xmlOutput.writeAttribute(ATTR_HEADER_NAME, name);
				}
				
				xmlOutput.writeCharacters(value);

				xmlOutput.writeEndElement();
			} else if (includeBlankCells)
				xmlOutput.writeEmptyElement(TAG_CELL);

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	@Override
	public void closeRow(XSSFRow row) throws IOException {

		try {
			xmlOutput.writeEndElement();

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	@Override
	public void closeSheet(XSSFSheet sheet) throws IOException {
		super.closeSheet(sheet);

		try {
			xmlOutput.writeEndElement();

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	@Override
	public void closeBook(XSSFWorkbook workbook) throws IOException {
		try {
			xmlOutput.writeEndElement();

			xmlOutput.writeEndDocument();

			output.println();

		} catch (XMLStreamException e) {
			throw new IOException(e);
		}
	}

	@Override
	public void setOutput(OutputStream out) {
		super.setOutput(out);

		this.xmlOutput = XMLStreamWriterFactory.create(out);

		if (this.indentOutput) {
			this.xmlOutput = new IndentingXMLStreamWriter(xmlOutput);
			// ((IndentingXMLStreamWriter)xmlOutput).setIndentStep("\t");
		}
	}

	@Override
	public void setArguments(NamedArgs args) {
		super.setArguments(args);

		this.includeBlankCells = args.contains(ARG_INCLUDE_BLANKS);

		this.excludeIndex = args.contains(ARG_XINDEX);
		this.excludeType = args.contains(ARG_XTYPE);
		this.includeHeaderNames = args.contains(ARG_INCLUDE_HEADNAMES);

		if (args.contains(ARG_INDENT)) {
			this.indentOutput = true;
			if (null != this.xmlOutput) // If output already set, reset so it
										// picks up indent wrapper.
				setOutput(output);
		}
	}
	@Override
	public void getArgumentHelp(PrintStream out) {
		
		out.println();
		
		out.println(getClass().getSimpleName() + " command line arguments:");
		out.println("-" + ARG_INDENT + "\t\tIndents the output XML.");
		out.println("\t\tUsually the XML is exported in a single line.");
		out.println("\t\tThis argument indents each tag to make it more human readable.");
		
		out.println();
		out.println("-" + ARG_XINDEX + "\tExclude the xxIndex attributes from the export.");
		out.println("\t\tUsually row and cells have their row and column indexes written in as attributes.");
		out.println("\t\tSpecifying this argument will supress those attributes from the output.");
		
		out.println();
		out.println("-" + ARG_XTYPE + "\t\tExclude the type attribute from the export.");
		out.println("\t\tUsually each element gets the data type of the value written into the 'type' attribute.");
		out.println("\t\tSpecifying this argument will supress that attribute from the output.");

		out.println();
		out.println("-" + ARG_INCLUDE_BLANKS + "\t\tIncludes blank cells in the output.");
		out.println("\t\tUsually cells containing nothing are not written to the export.");
		out.println("\t\tSpecifying this argument will include empty, leaf nodes where blank cells are found.");

		out.println();
		out.println("-" + ARG_INCLUDE_HEADNAMES + "\tIncludes " + ATTR_HEADER_NAME + " attribute in each cell element.");
		out.println("\t\tSpecifying this argument will write a " + ATTR_HEADER_NAME + " attribute containing the header name in the aligning column.");
		

	}



}
