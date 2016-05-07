package ca.mcgilleus.budgetbuilder.controller;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.LinkedList;
import java.util.Queue;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;

/**
 * Kareem Halabi
 * 260 616 162
 */

public class PortfolioCreator {
	
	private static Queue<IndexedColors> allColors;
	private static Queue<IndexedColors> availableTabColors;
	
	public static Queue<IndexedColors> getAllColors() {
		return allColors;
	}

	public static void setAllColors(Queue<IndexedColors> allColors) {
		PortfolioCreator.allColors = allColors;
		recreateColorQueue();
	}
	
	private static void recreateColorQueue() {
		availableTabColors = new LinkedList<IndexedColors>(allColors);
	}
	
	/**
	 * Creates an EUS portfolio from a directory
	 * @param f The root directory of the portfolio
	 * @param budget Budget to add portfolio to  
	 */
	public static void createPortfolio(File f, EUSBudget budget) {
		Portfolio p = new Portfolio(f.getName(), budget);

System.out.println("Created portfolio: " + p.getName());

		budget.getWorkbook().createSheet(p.getName());
		File[] portfolioFiles = f.listFiles();

		for (File pF : portfolioFiles) {
			try {
				CommitteeCreator.createCommitteeBudget(availableTabColors.peek(), p, pF);
			} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		availableTabColors.poll();

		if (availableTabColors.isEmpty())
			recreateColorQueue();

		PortfolioCreator.createPortfolioOverView(p);
	}

	//TODO styling
	public static void createPortfolioOverView(Portfolio p) {
		
		XSSFWorkbook wb = p.getEUSBudget().getWorkbook();
		
		XSSFSheet overviewSheet = wb.getSheet(p.getName());	
		XSSFRow header = overviewSheet.createRow(0);
		
		XSSFCell title = header.createCell(0, Cell.CELL_TYPE_STRING);
		title.setCellValue("Function");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		title.setCellValue("Committee");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		
		Calendar c = Calendar.getInstance();
		c.setTime(p.getEUSBudget().getYear());
		String budgetYear = c.get(Calendar.YEAR) + "-" + (c.get(Calendar.YEAR) +1);
		
		title.setCellValue("Budget " + budgetYear);
		
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
		
		//-- Write totals --//
		XSSFRow totals = overviewSheet.createRow(overviewSheet.getLastRowNum()+2);
		XSSFCell totalLabel = totals.createCell(1, Cell.CELL_TYPE_STRING);
		totalLabel.setCellValue("Totals");
		
		for(int j = totals.getFirstCellNum() + 1;
				j < overviewSheet.getRow(0).getLastCellNum(); j++) {
			XSSFCell colTotal = totals.createCell(j, Cell.CELL_TYPE_FORMULA);
			CellAddress ref = colTotal.getAddress();
			String sumFormula = "SUM(" + CellReference.convertNumToColString(ref.getColumn()) + "2:" +
					CellReference.convertNumToColString(ref.getColumn()) + (ref.getRow()-1) + ")";
			colTotal.setCellFormula(sumFormula);
		}
		
System.out.println("\t Overview done!");		
	}

	
}
