package com.smeup.test;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;


public class PrintToExcel {

	private static Logger logger = LoggerFactory.getLogger(PrintToExcel.class);

	public static void main(String[] args) throws ParseException, IOException, Exception {
		logger.info("Running Object Collection demo");
		List<String> strings = new ArrayList<String>();
		strings.add("Test");
		strings.add("Test2");
		strings.add("Test3");
		strings.add("Test4");
		int i=0;

		String tempPath = "src/main/resources/excel/grid_template.xlsx";
		FileInputStream is = new FileInputStream(new File(tempPath));
		try (OutputStream os = new FileOutputStream("src/main/resources/excel/grid_output.xlsx")) {
			Context context = new Context();
			context.putVar("strings", strings);
			context.putVar("i", i);
			JxlsHelper.getInstance().processTemplate(is, os, context);
		}
	}

}
