package ca.mcgilleus.test;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTest {
	public static void main(String[] args) {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFCreationHelper createHelper = wb.getCreationHelper();
		String s = WorkbookUtil.createSafeSheetName("sheet1");
		XSSFSheet sheet = wb.createSheet(s);
		sheet.setTabColor(IndexedColors.RED.getIndex());
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
	    cell.setCellValue(createHelper.createRichTextString("HelloWorld"));
	    
	    try {
			FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    
	    try {
			XSSFWorkbook w = (XSSFWorkbook) WorkbookFactory.create(new File("workbook.xlsx"));
			XSSFSheet sheet1 = w.getSheet("sheet1");
			XSSFRow row1 = sheet1.getRow(0);
			XSSFCell cell1 = row1.getCell(0);
			System.out.println(cell1.getRichStringCellValue());
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    
	}
}
