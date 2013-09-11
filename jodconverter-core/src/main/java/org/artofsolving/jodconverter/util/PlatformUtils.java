//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter.util;

public class PlatformUtils {

	private static final String OS_NAME = System.getProperty("os.name").toLowerCase();

	private PlatformUtils() {
		throw new AssertionError("utility class must not be instantiated");
	}

	public static boolean isLinux() {
		return OS_NAME.startsWith("linux");
	}

	public static boolean isMac() {
		return OS_NAME.startsWith("mac");
	}

	public static boolean isWindows() {
		return OS_NAME.startsWith("windows");
	}

	/**
	 * @author : dmotta
	 * @date : Jul 19, 2013
	 * @description : TODO(dmotta) : insert description
	 * @return : int TODO(dmotta)
	 */
	public static int firstNumericPosition(String cellName) {
		int position = 0;
		for (int y = 0; y < cellName.length(); y++) {
			char result = cellName.charAt(y);
			if (isInteger(result)) {
				position = y;
				System.out.println("position: " + position);
				break;
			}
		}
		return position;
	}

	/**
	 * @author : dmotta
	 * @date : Jul 19, 2013
	 * @description : TODO(dmotta) : insert description
	 * @return : boolean TODO(dmotta)
	 */
	public static boolean isInteger(char input) {
		try {
			Integer.parseInt(input + "");
			return true;
		}
		catch (Exception e) {
			return false;
		}
	}

}
