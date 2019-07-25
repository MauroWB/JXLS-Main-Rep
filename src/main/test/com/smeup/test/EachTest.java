package com.smeup.test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

public class EachTest {

	public static void main(String[] args) throws IOException {
		final InputStream in = new FileInputStream("src/main/resources/test/each_template.xlsx");
		final OutputStream out = new FileOutputStream("src/main/resources/test/each_output.xlsx");
		
		Context context = new Context();
		ArrayList<Object> arr = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			arr.add("T"+i);
		}
		context.putVar("arr", arr);
		JxlsHelper.getInstance().processTemplate(in, out, context);
		
		in.close();
		out.close();
	}

}
