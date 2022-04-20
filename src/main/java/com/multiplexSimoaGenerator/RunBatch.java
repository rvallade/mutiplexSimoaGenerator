package com.multiplexSimoaGenerator;

import java.io.IOException;

public class RunBatch {
	public RunBatch() {
	}

	public static void main(String[] args) throws IOException {
		Generator generator = new Generator();

		try {
			generator.execute();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}