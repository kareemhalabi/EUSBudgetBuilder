package ca.mcgilleus.budgetbuilder.controller;

import java.io.File;
import java.io.IOException;
import java.util.LinkedList;
import java.util.Queue;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Styles;

/**
 * Kareem Halabi
 * 260 616 162
 */

public class PortfolioCreator {
	
	private static Queue<IndexedColors> allColors;
	private static Queue<IndexedColors> availableTabColors;
	private static IndexedColors currentColor;
	
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

System.out.println("Compiling portfolio: " + p.getName());

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

		currentColor = availableTabColors.poll();

		if (availableTabColors.isEmpty())
			recreateColorQueue();

		createPortfolioOverView(p);
	}

	/**
	 * Creates Portfolio Overview page showing amounts requested by portfolio's committees
	 * @param p The porfilo overview to create
	 */
	private static void createPortfolioOverView(Portfolio p) {
		
		XSSFWorkbook wb = p.getEUSBudget().getWorkbook();
		
		XSSFSheet overviewSheet = wb.getSheet(p.getName());	
		XSSFRow header = overviewSheet.createRow(0);
		
		XSSFCell title = header.createCell(0, Cell.CELL_TYPE_STRING);
		title.setCellValue("Function");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		title.setCellValue("Committee");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		
		title.setCellValue(p.getEUSBudget().getBudgetYear() + " Budget");
		
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
		
		stylePortfolioOverview(overviewSheet);
		
System.out.println("\t Overview done!");		
	}

	private static void stylePortfolioOverview(XSSFSheet overviewSheet) {
		
		XSSFRow header = overviewSheet.getRow(0);
		
		//Fix Column Widths
		for(int i = 0; i < header.getLastCellNum(); i++) {
			overviewSheet.autoSizeColumn(i);
			int adjustedWidth = overviewSheet.getColumnWidth(i) + 256*3;
			
			if(adjustedWidth > 5000)
				adjustedWidth = 5000;
			
			overviewSheet.setColumnWidth(i, adjustedWidth);
		}
		
		//Apply Header styles
		for(Cell cell : header) {
			cell.setCellStyle(Styles.HEADER_STYLE);
		}
		
		//Apply Data styles
		for(int j = 1; j < overviewSheet.getLastRowNum()-1; j++) {
			XSSFRow dataRow = overviewSheet.getRow(j);
			
			CellStyle customPortfolioStyle = EUSBudgetBuilder.getWorkbook().createCellStyle();
			customPortfolioStyle.cloneStyleFrom(Styles.PORTFOLIO_LABEL_STYLE);
			XSSFFont customPortfolioFont = EUSBudgetBuilder.getWorkbook().createFont();
			customPortfolioFont.setFontHeightInPoints((short)10);
			customPortfolioFont.setFontName("Arial");
			customPortfolioFont.setColor(currentColor.getIndex());
			customPortfolioFont.setBold(false);
			customPortfolioFont.setItalic(false);
			
			customPortfolioStyle.setFont(customPortfolioFont);
			
			dataRow.getCell(0).setCellStyle(customPortfolioStyle);
			
			dataRow.getCell(1).setCellStyle(Styles.COMMITTEE_LABEL_STYLE);
			
			for(int k = 2; k < dataRow.getLastCellNum(); k++)
				dataRow.getCell(k).setCellStyle(Styles.CURRENCY_CELL_STYLE);
		}
		
		//Apply Total styles
		XSSFRow totals= overviewSheet.getRow(overviewSheet.getLastRowNum());
		XSSFCell totalLabel = totals.getCell(1);
		totalLabel.setCellStyle(Styles.TOTAL_LABEL_STYLE);
		
		for(int l = 2; l < totals.getLastCellNum(); l++) 
			totals.getCell(l).setCellStyle(Styles.TOTAL_CELL_STYLE);
	}
}
