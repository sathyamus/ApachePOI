/**
 * 
 */
package com.snsystems.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

/**
 * @author SN
 *
 */
public class POIUtil {
	
	private Log log = LogFactory.getLog(POIUtil.class);
	private static final String XLS = ".xls";
	private static final String XLSX = ".xlsx";
	
	public Workbook createWorkbook(String fileNameWithExtn, boolean isXlsx) {
		// New Workbook
	    Workbook workbook = new HSSFWorkbook();
	    FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(fileNameWithExtn);
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	    return workbook;
	}

    public Sheet createSheet(String sheetName, Workbook workbook) {
    	// New Sheet
    	
        // Note that sheet name is Excel must not exceed 31 characters
        // and must not contain any of the any of the following characters:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])
    	
    	if (null != workbook && null != sheetName) {
    		return workbook.createSheet(sheetName);
    	}
        return null;
    }

    public String createSafeSheetName(String sheetName) {
        // for a safe way to create valid names, this utility replaces invalid characters with a space (' ')
    	return WorkbookUtil.createSafeSheetName(sheetName);
    }

//
//Creating Cells
//    Workbook wb = new HSSFWorkbook();
//    //Workbook wb = new XSSFWorkbook();
//    CreationHelper createHelper = wb.getCreationHelper();
//    Sheet sheet = wb.createSheet("new sheet");
//
//    // Create a row and put some cells in it. Rows are 0 based.
//    Row row = sheet.createRow((short)0);
//    // Create a cell and put a value in it.
//    Cell cell = row.createCell(0);
//    cell.setCellValue(1);
//
//    // Or do it on one line.
//    row.createCell(1).setCellValue(1.2);
//    row.createCell(2).setCellValue(
//         createHelper.createRichTextString("This is a string"));
//    row.createCell(3).setCellValue(true);
//
//    // Write the output to a file
//    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
//    wb.write(fileOut);
//    fileOut.close();
//                    
//Creating Date Cells
//    Workbook wb = new HSSFWorkbook();
//    //Workbook wb = new XSSFWorkbook();
//    CreationHelper createHelper = wb.getCreationHelper();
//    Sheet sheet = wb.createSheet("new sheet");
//
//    // Create a row and put some cells in it. Rows are 0 based.
//    Row row = sheet.createRow(0);
//
//    // Create a cell and put a date value in it.  The first cell is not styled
//    // as a date.
//    Cell cell = row.createCell(0);
//    cell.setCellValue(new Date());
//
//    // we style the second cell as a date (and time).  It is important to
//    // create a new cell style from the workbook otherwise you can end up
//    // modifying the built in style and effecting not only this cell but other cells.
//    CellStyle cellStyle = wb.createCellStyle();
//    cellStyle.setDataFormat(
//        createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
//    cell = row.createCell(1);
//    cell.setCellValue(new Date());
//    cell.setCellStyle(cellStyle);
//
//    //you can also set date as java.util.Calendar
//    cell = row.createCell(2);
//    cell.setCellValue(Calendar.getInstance());
//    cell.setCellStyle(cellStyle);
//
//    // Write the output to a file
//    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
//    wb.write(fileOut);
//    fileOut.close();
//                    
//Working with different types of cells
//    Workbook wb = new HSSFWorkbook();
//    Sheet sheet = wb.createSheet("new sheet");
//    Row row = sheet.createRow((short)2);
//    row.createCell(0).setCellValue(1.1);
//    row.createCell(1).setCellValue(new Date());
//    row.createCell(2).setCellValue(Calendar.getInstance());
//    row.createCell(3).setCellValue("a string");
//    row.createCell(4).setCellValue(true);
//    row.createCell(5).setCellType(Cell.CELL_TYPE_ERROR);
//
//    // Write the output to a file
//    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
//    wb.write(fileOut);
//    fileOut.close();
}
