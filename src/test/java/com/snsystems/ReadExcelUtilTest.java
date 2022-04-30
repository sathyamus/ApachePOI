package com.snsystems;

import net.sf.jett.transform.ExcelTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertTrue;

public class ReadExcelUtilTest {

	@Test
	public void test() throws Exception {

		String path = Paths.get("src/main/resources/ExcelSheet.xls").toUri().getPath();
		System.out.println(path);

		int rowsCount = ReadExcelUtil.read(path);
		System.out.println(rowsCount);
		assertTrue("Row Count: ", rowsCount != 0);

		Map<String, Object> beans = new HashMap<String, Object>();
		beans.put("var", "Hello");
		beans.put("var2", "World");
		ExcelTransformer transformer = new ExcelTransformer();
		try
		{
		   transformer.transform("template.xlsx", "result.xlsx", beans);
		}
		catch (IOException e)
		{
		   System.err.println("I/O error occurred: " + e.getMessage());
		}
		catch (InvalidFormatException e)
		{
		   System.err.println("Spreadsheet was in invalid format: " + e.getMessage());
		}
		
		// org.apache.poi.ss.util.SheetUtil;
		// net.sf.jett.util.SheetUtil util = new SheetUtil();
		 	
	}

}
