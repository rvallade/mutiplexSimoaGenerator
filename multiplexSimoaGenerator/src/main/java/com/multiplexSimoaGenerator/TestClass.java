package com.multiplexSimoaGenerator;

public class TestClass {
	public static void main(String[] args) {
		System.out.println(handleDoubleQuotes("abc,abc,abc"));
		System.out.println(handleDoubleQuotes("abc,\"1,234.12\",abc"));
		System.out.println(handleDoubleQuotes("abc,\"1,234\",\"1,234.12\",abc"));
		System.out.println(handleDoubleQuotes("\"1,234.12\",abc,\"1,234\",abc"));
		System.out.println(handleDoubleQuotes("abc,\"1,234.12\""));
		System.out.println(handleDoubleQuotes("abc,\"1,234.12\",abc,\"2,345\""));
		System.out.println(handleDoubleQuotes("\"2,345\",abc,\"1,234.12\""));
	}
	
	public static String handleDoubleQuotes(String line) {
		//System.out.println(line);
		if (line.contains("\"")) {
			// sometimes a number is in "" with a comma to separate thousands, like
			// ...AP,Complete,21,1.962676,"1,056.00",-,1,...
			// if we split using the comma only it produces a bug
			/*if (line.startsWith("\"")) {
				System.out.println("\tLine starts with \"");
				position = 0;
			}*/
			String[] datas = line.split("\"");
			for (int position = 1 ; position < datas.length ; ) {
				//System.out.println("\tposition=" + position);
				//System.out.println("\tdatas[position] = " + datas[position]);
				//System.out.println("\tReplacement result: " + line.replace("\"" + datas[position] + "\"", datas[position].replace(",", "")));
				line = line.replace("\"" + datas[position] + "\"", datas[position].replace(",", ""));
				//System.out.println("\tline=" + line);
				position = position + 2;
			}
		}
		return line;
	}
}