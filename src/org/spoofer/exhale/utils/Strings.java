package org.spoofer.exhale.utils;

public class Strings
{

	private Strings(){}
	
	public static boolean isEmpty(String s)
	{
		return null == s || s.trim().length() < 1;
	}
	
	public static String capitalise(String s)
	{
		if (isEmpty(s))
			return s;
		
		s = s.trim();
		StringBuilder sb = new StringBuilder();
		sb.append(Character.toTitleCase(s.charAt(0)));
		if (s.length() > 1)
			sb.append(s.substring(1).toLowerCase());
		
		return sb.toString();
	}
	
}
