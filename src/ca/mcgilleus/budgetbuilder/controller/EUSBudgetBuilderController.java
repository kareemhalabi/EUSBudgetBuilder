package ca.mcgilleus.budgetbuilder.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.Queue;
import java.util.Stack;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Util;

/**
 * Kareem Halabi
 * 260 616 162
 */

public class EUSBudgetBuilderController {
	
	static XSSFWorkbook wb;

	public static void main(String[] args) {
		
		File root = new File(System.getProperty("user.dir") + File.separator + "W2017 Budget");
		EUSBudget budget = createBudget();
		wb = budget.getBudget();
		
		Queue<IndexedColors> availableTabColors = new LinkedList<IndexedColors>();
		recreateColorQueue(availableTabColors);
		
		for(File f : root.listFiles()) {
			Portfolio p = new Portfolio(f.getName(), budget);
System.out.println("Created portfolio: " + p.getName());
			wb.createSheet(p.getName());
			File[] portfolioFiles = f.listFiles();
			
			for(File pF: portfolioFiles) {
				try {
					createCommitteeBudget(availableTabColors.peek(), p, pF);	
				} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			
			availableTabColors.poll();
			
			if(availableTabColors.isEmpty())
				recreateColorQueue(availableTabColors);
			
			createPortfolioOverView(p);

		}
		
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(root.getName() + ".xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	public static void recreateColorQueue(Queue<IndexedColors> colors) {
		colors.add(IndexedColors.MAROON);
		colors.add(IndexedColors.LIGHT_ORANGE);
		colors.add(IndexedColors.LIGHT_YELLOW);
		colors.add(IndexedColors.LIGHT_GREEN);
		colors.add(IndexedColors.LIGHT_BLUE);
		colors.add(IndexedColors.PLUM);
		colors.add(IndexedColors.GREY_40_PERCENT);
		
	}

	public static void createCommitteeBudget(IndexedColors color, Portfolio p, File pF)
			throws IOException, InvalidFormatException {
		
		XSSFWorkbook pWB = (XSSFWorkbook) WorkbookFactory.create(pF);
		XSSFSheet pSheet = pWB.getSheetAt(0);
		
		Name name = pWB.getName("COMM_NAME");
		CellReference nameRef = new CellReference(name.getRefersToFormula());
		XSSFRow nameRow = pSheet.getRow(nameRef.getRow());
		XSSFCell nameCell = nameRow.getCell(nameRef.getCol());
		String sheetName = nameCell.getStringCellValue();
		
		//Get full name if abbreviated unavailable
		if(sheetName == null || sheetName.trim().length() == 0 ) {
			nameRow = pSheet.getRow(nameRef.getRow() - 1);
			nameCell = nameRow.getCell(nameRef.getCol());
			sheetName = nameCell.getStringCellValue();
		}
		
		//So that AMT reference has correct sheet name
		pWB.setSheetName(0, sheetName);
		
		Name amt = pWB.getName("AMT");
		
		CommitteeBudget committeeBudget = new CommitteeBudget(sheetName, amt.getRefersToFormula(), p);
		
		XSSFSheet bSheet = wb.createSheet(sheetName);
		Util.cloneSheet(pSheet, bSheet);
		
		bSheet.setTabColor(color.getIndex());
System.out.println("\t Created committee: " + sheetName);
	}

	//TODO styling
	public static void createPortfolioOverView(Portfolio p) {
		XSSFSheet overviewSheet = wb.getSheet(p.getName());	
		XSSFRow header = overviewSheet.createRow(0);
		header.setHeightInPoints(45);
		
		XSSFCell title = header.createCell(0, Cell.CELL_TYPE_STRING);
		title.setCellValue("Function");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		title.setCellValue("Committee");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		title.setCellValue("Budget");
		
		for(int i = 0; i < p.getCommitteeBudgets().size(); i++) {
			CommitteeBudget committee = p.getCommitteeBudget(i);
			XSSFRow committeeRow = overviewSheet.createRow(i+1);
			
			XSSFCell committeeFunction = committeeRow.createCell(0, Cell.CELL_TYPE_STRING);
			committeeFunction.setCellValue(p.getName());
			
			XSSFCell committeeName = committeeRow.createCell(committeeRow.getLastCellNum(), Cell.CELL_TYPE_STRING);
			committeeName.setCellValue(committee.getName());
			
			XSSFCell committeeAmt = committeeRow.createCell(committeeRow.getLastCellNum(), Cell.CELL_TYPE_FORMULA);
			committeeAmt.setCellFormula(committee.getAmtRequestedRef());
		}
System.out.println("\t Overview done!");		
	}

	public static EUSBudget createBudget() {
		Calendar c = Calendar.getInstance();
		Date budgetYear;
		if(c.get(Calendar.MONTH) <= Calendar.FEBRUARY)
			c.add(Calendar.YEAR, -1);
		
		budgetYear = c.getTime();
		
		return new EUSBudget(budgetYear);
	}
	
}
