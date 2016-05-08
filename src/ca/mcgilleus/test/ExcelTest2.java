package ca.mcgilleus.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.util.Cloner;

/**
 * Kareem Halabi
 * 260 616 162
 */

public class ExcelTest2 {
	public static void main(String[] args) {
		
		try {
			XSSFWorkbook wb;
			wb = (XSSFWorkbook) WorkbookFactory.create(new File("Test.xlsx"));
//			CellReference cr = new CellReference("B37");
			XSSFSheet sheet = wb.getSheet("Template");
			XSSFSheet duplicate = wb.createSheet();
			Cloner.cloneSheet(sheet, duplicate);
			Name n = wb.getName("AMT");
			System.out.println(n.getRefersToFormula());
			Name m = wb.getNameAt(1);
			System.out.println(m.getRefersToFormula());
			FileOutputStream fileOut = new FileOutputStream("Test2.xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e1) {
			e1.printStackTrace();
		}
	    
	}
}
