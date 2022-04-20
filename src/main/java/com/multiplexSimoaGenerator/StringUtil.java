package com.multiplexSimoaGenerator;

public class StringUtil {
	public static boolean isEmpty(String value) {
		return (value == null) || (value.trim().isEmpty());
	}
	
	public static boolean isSameSample(String name1, String name2) {
		return isSameSampleNameForBothMode(name1, name2) || isSampleNamesSuffixedWithDigits(name1, name2);
	}
	
	public static String getCommonSampleName(String name, boolean sampleNameUsedAsIsInDuplicate) {
		return sampleNameUsedAsIsInDuplicate ? name : name.substring(0, name.length() - 2);
	}
	
	private static boolean isSampleNamesSuffixedWithDigits(String name1, String name2) {
		return name1.substring(0, name1.length() - 2).equals(name2.substring(0, name2.length() - 2));
	}

	public static boolean isSameSampleNameForBothMode(String name1, String name2) {
		// In the first versions of Simoa we would see SAMPLE_NAME01 and SAMPLE_NAME02, that is why we would remove the last 2 digits
		// Now their software keeps the sample id for both, no need to remove anything
		return name1.equals(name2);
	}
}