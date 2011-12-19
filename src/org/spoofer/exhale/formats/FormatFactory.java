package org.spoofer.exhale.formats;

import java.io.IOException;
import java.util.Properties;

public class FormatFactory {

	private static final String FORMAT_PROPERTIES = "formats.properties";

	@SuppressWarnings("unchecked")
	public static ExhaleFormat loadFormat(String format) throws FormatException {

		Properties formats = getProperties();

		if (!formats.containsKey(format.toLowerCase()))
			throw new FormatException("Unknown format '" + format + " '");

		String formatClassName = formats.getProperty(format);
		Class<ExhaleFormat> formatClass = null;

		try {
			formatClass = (Class<ExhaleFormat>) Class.forName(formatClassName);

		} catch (ClassNotFoundException e) {
			throw new FormatException("Failed to find a class for format " + formatClassName);
		}
		
		if (!ExhaleFormat.class.isAssignableFrom(formatClass))
			throw new FormatException(formatClass.getName()
					+ " does not support the " + ExhaleFormat.class.getName()
					+ " interface");

		ExhaleFormat newFormat;
		try {
			newFormat = formatClass.newInstance();

		} catch (Exception e) {
			throw new FormatException("failed to load format "
					+ formatClass.getName(), e);
		}

		return newFormat;
	}

	public static String[] getFormats() {

		Properties props = getProperties();

		String[] names = new String[props.size()];
		int index = 0;
		for (String name : props.stringPropertyNames())
			names[index++] = name; 

		return names;
	}

	protected static Properties getProperties() {
		Properties props = new Properties();
		try {
			props.load(FormatFactory.class.getResourceAsStream(FORMAT_PROPERTIES));
		
		} catch (IOException e) {
			throw new IllegalStateException("Failed to open formats property file: " + FORMAT_PROPERTIES);
		}

		return props;
	}


}
