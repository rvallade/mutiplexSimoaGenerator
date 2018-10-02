package com.multiplexSimoaGenerator;

public class StringUtil {
	public static boolean isEmpty(String value) {
		return (value == null) || (value.trim().isEmpty());
	}
	
	public static boolean isSameSample(String name1, String name2) {
		/*boolean result = false;
		if (name1.length() == name2.length()) {
			int totalSimilarity = 0;
			for (int i = 0 ; i < name1.length() ; i++) {
				if (name1.charAt(i) == name2.charAt(i)) {
					totalSimilarity++;
				} else {
					break;
				}
			}
			result = totalSimilarity >= (name1.length() - 2);
		}
		return result;*/
		return getCommonSampleName(name1).equals(getCommonSampleName(name2));
	}
	
	/*public static String getSampleName(String sampleID) {
		int lastLetterBeforeNumbers = 0;
		for (int i = sampleID.length()-1 ; i > 0 ; i--) {
			if (!Character.isDigit(sampleID.charAt(i))) {
				lastLetterBeforeNumbers = i;
				break;
			}
		}
		return sampleID.substring(0, lastLetterBeforeNumbers + 1);
	}*/
	
	public static String getCommonSampleName(String name1) {
		/*String commonName = "";
		if (name2 == null) {
			commonName = name1;
		} else {
			for (int i = 0 ; i < name1.length() ; i++) {
				if (name1.charAt(i) == name2.charAt(i)) {
					commonName += name1.charAt(i);
				} else {
					break;
				}
			}
		}
		return commonName;*/
		return name1.substring(0, name1.length()-2);
	}
}