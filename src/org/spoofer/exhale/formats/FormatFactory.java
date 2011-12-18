package org.spoofer.exhale.formats;

import static org.spoofer.exhale.utils.Strings.isEmpty;

import org.spoofer.exhale.utils.Strings;

public class FormatFactory
{

	public static final String DEFAULT_FORMAT_PACKAGE = Csv.class.getPackage().getName();
	
	private static final String PACKAGE_NAME_DELIMITER = ";";


	
	
	@SuppressWarnings("unchecked")
	public static ExhaleFormat loadFormat(String format) throws FormatException
	{
		String[] packages = getFormatPackages();
		
		String formatClassName = Strings.capitalise(format);
		Class<ExhaleFormat> formatClass = null;
		
		for (String packName : packages)
		{
			StringBuilder cls = new StringBuilder(packName);
			if (!packName.endsWith("."))
				cls.append(".");
			cls.append(formatClassName);
			
			try {
				formatClass = (Class<ExhaleFormat>) Class.forName(cls.toString());
			
			} catch (ClassNotFoundException e) {
				formatClass = null;
			}
		
			if (null != formatClass)
				break;
		}

		if (null == formatClass)
			throw new FormatException("Failed to find a class for format " + formatClassName);
		
		if (!ExhaleFormat.class.isAssignableFrom(formatClass))
			throw new FormatException(formatClass.getName() + " does not support the " + ExhaleFormat.class.getName() + " interface");
	
		ExhaleFormat newFormat;
		try {
			newFormat = formatClass.newInstance();
			
		} catch (Exception e) {
			throw new FormatException("failed to load format " + formatClass.getName(), e);
		}
		
		return newFormat;
	}

	public static String[] getFormats()
	{
		Class<ExhaleFormat>[] classes = (Class<ExhaleFormat>[]) ExhaleFormat.class.getClasses();
		String[] names = new String[classes.length];
		
		int index = 0;
		for (Class<ExhaleFormat> cls : classes)
			names[index++] = cls.getSimpleName().toLowerCase();
		
		return names;
	}
	public static String[] getFormatPackages()
	{
		String packageName = System.getProperty(DEFAULT_FORMAT_PACKAGE);
		if (isEmpty(packageName))
			packageName = DEFAULT_FORMAT_PACKAGE;
		
		return packageName.contains(PACKAGE_NAME_DELIMITER) ? packageName.split(PACKAGE_NAME_DELIMITER) : new String[]{packageName.trim()};
	
	}
}
