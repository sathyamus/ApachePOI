package com.snsystems;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;


public class ReadExcelUtil {

	public static int read(String fileName) throws Exception {
		
		FileInputStream file = new FileInputStream(new File(fileName));
		
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		
		HSSFSheet sheet = workbook.getSheetAt(0);

		//Iterate through each rows from first sheet
        Iterator<Row> rowIterator = sheet.iterator();
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
             
            //For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()) {
                 
                Cell cell = cellIterator.next();

                if ( CellType.STRING.equals(cell.getCellTypeEnum()) ) {
                    System.out.println(cell.getStringCellValue());
                }
                if ( CellType.NUMERIC.equals(cell.getCellTypeEnum()) ) {
                    System.out.println(cell.getNumericCellValue());
                }
            }
        }
            
		return sheet.getLastRowNum();

	}

}
