package org.spoofer.exhale.utils;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * A command line arguments parcer takes the original args string array from the
 * <code>main</code> method and parses it into a series of named arguments.
 * 
 * Named arguments are expected to use the '-' switch character to signal an
 * argument name. The name can be, optionally, followed by a value for that
 * argument, separated by a white space. (Arguments names can not contain
 * spaces) The value of the argument will include all the argument string up to
 * but excluding the following name switch ('-') or the end of the argument
 * string. Arguments do not have to contain a value, if the character imeedeatly
 * following the white space after the previous name is a switch, then the
 * argument will exist, but with no value. In this case, it's existence can be
 * checked with <code>contains(name)</code>, but requesting its value with
 * <code>getArgument(name)</code> will return null.
 * 
 * e.g. 1 sample.MyClass -argumentOne 1 -argumentTwo 2 -argumentThree 3
 * -argumentFour Four named arguments, named 'argumentOne', 'argumentTwo' ...
 * 'argumentFour' Each with a value of 1,2 and 3 respectivly, with the exception
 * of argumentFour, which has no value.
 * 
 * 
 * Unnamed Argument The initial part of the argument string does not have to be
 * named. This is known, internally, as the UNNAMED argument. It is accessed
 * publicly by using a null for the argument name.
 * 
 * e.g. 2 sample.MyClass argumenttest -argumentOne 1 -argumentTwo 2
 * -argumentThree 3 -argumentFour Would have the same named arguments as the
 * previous example, but an additional unnamed argument, with a value of
 * 'argumenttest'. This unnamed argument can be used like any other named
 * argument, by simply replacing the argument name with a null. e.g.
 * contains(null); is valid and in the first example this would return false, in
 * the second, true. getArgument(null), would return null in eg1, and
 * "argumenttest" in eg2
 * 
 * 
 * @author robgilham
 * 
 */
public class NamedArgs {

	/**
	 * The SWITCH is the character used in the Arguments as a name indicator.
	 * Argument names should be placed directly after a switch, with no white
	 * space.
	 */
	public static final String SWITCH = "-";

	/**
	 * The name of the unname argument (Oxymoron or what!)
	 */
	private static final String UNNAMED_ARG = "__UNNAMED";

	/**
	 * The main store for the named arguments
	 */
	private final Map<String, String> arguments;

	/**
	 * Create a new Named arguments, using the original main string array of
	 * arguments.
	 * 
	 * @param args
	 *            An array of arguments as supplied by the JRE at start up.
	 */
	public NamedArgs(String[] args) {
		super();

		arguments = parseArgs(args);

	}

	/**
	 * Gets any value that the named argument has. If the named argument, has
	 * any characters following it, these are counted as its value.
	 * 
	 * @param name
	 *            the name of the argument. Null can be used to retrieve any
	 *            value stated before any named arguments in the command line.
	 * 
	 * @return the value of that argument or null if the argument has no value
	 *         or does not exist.
	 */
	public String getArgument(String name) {
		if (null == name)
			name = UNNAMED_ARG;

		return arguments.get(name);
	}

	/**
	 * Checks if the given name exists as a named argument. If the argument name
	 * was present in the command line, this will be true, regardless if the the
	 * named argument has a value or not.
	 * 
	 * @param name
	 *            the name of the argument to check
	 * @return true if the name matches a named argument, false otherwise.
	 */
	public boolean contains(String name) {
		if (null == name || name.trim().length() < 1)
			name = UNNAMED_ARG;
		return arguments.containsKey(name);
	}

	/**
	 * Checks if the given name both exists (calls contains) AND has a value. If
	 * the given name exists as a named argument and the argument has a value
	 * that is not null, not empty and has more than white space in it.
	 * 
	 * @param name
	 *            the name of the argument to check
	 * @return true if the named argument exists and has a value, false
	 *         otherwise
	 */
	public boolean hasValue(String name) {
		String value = this.getArgument(name);
		return null != value && value.trim().length() > 0;
	}

	/**
	 * Gets the total number of named arguments. This will include any unnamed
	 * argument if it is present.
	 * 
	 * @return to total number of arguments
	 */
	public int getCount() {
		return arguments.size();
	}

	/**
	 * Gets a Set of all the names of the named arguments
	 * 
	 * @return all the names of the arguments.
	 */
	public Set<String> getArgumentNames() {
		Set<String> result = arguments.keySet();
		if (result.contains(UNNAMED_ARG))
			result.remove(UNNAMED_ARG);

		return arguments.keySet();
	}

	/**
	 * Parse the given argument of strings into the Map of named arguments.
	 * 
	 * @param args
	 *            the original main args string array
	 * @return
	 */
	private Map<String, String> parseArgs(String[] args) {
		StringBuffer oneArg = new StringBuffer();
		for (String arg : args) // re-assemble the array into a single line
								// again.
		{
			if (oneArg.length() > 0)
				oneArg.append(' ');
			oneArg.append(arg);
		}

		String[] splitArgs = oneArg.toString().split(SWITCH); // Break it down
																// by the name
																// switches
		Map<String, String> newArgs = new HashMap<String, String>();

		if (splitArgs.length > 0) {
			String unnamed = splitArgs[0].trim();
			if (unnamed.length() > 0)
				newArgs.put(UNNAMED_ARG, unnamed);

			for (int i = 1; i < splitArgs.length; i++) {
				int iPos = splitArgs[i].indexOf(' '); // Move to next space
				String name = iPos > 0
						? splitArgs[i].substring(0, iPos).trim()
						: splitArgs[i].trim();
				String value = iPos > 0 ? splitArgs[i].substring(iPos + 1)
						.trim() : null;
				newArgs.put(name, value);
			}
		}
		return newArgs;
	}

	public Map<String, String> getArgMap() {
		Map<String, String> newArgs = new HashMap<String, String>();
		newArgs.putAll(arguments);
		if (this.contains(null)) {
			newArgs.put(null, getArgument(null));
			newArgs.remove(UNNAMED_ARG);
		}

		return newArgs;

	}

	public String getArgumentString() {
		StringBuilder sb = new StringBuilder();
		if (arguments.containsKey(UNNAMED_ARG))
			sb.append(arguments.get(UNNAMED_ARG));

		for (String arg : arguments.keySet())
			if (!arg.equals(UNNAMED_ARG)) {

				if (sb.length() > 0)
					sb.append(' ');

				sb.append(SWITCH).append(arg);

				if (this.hasValue(arg))
					sb.append(' ').append(arguments.get(arg));
			}

		return sb.toString();
	}

}
