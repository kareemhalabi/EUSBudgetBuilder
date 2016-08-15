/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.controller;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Styles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.util.LinkedList;
import java.util.Queue;

import static ca.mcgilleus.budgetbuilder.controller.BudgetBuilder.*;

public class PortfolioCreator {
	
	private static Queue<IndexedColors> allColors;
	private static Queue<IndexedColors> availableTabColors;
	private static IndexedColors currentColor;
	private static Portfolio p;
	
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
	 * @param portfolioDirectory The root directory of the portfolio
	 * @param budget Budget to add portfolio to  
	 */
	public static void createPortfolio(File portfolioDirectory, EUSBudget budget) {
		p = new Portfolio(portfolioDirectory.getName(), budget);

		buildTask.updateBuildMessage("Compiling portfolio: " + p.getName());

		budget.getWorkbook().createSheet(p.getName());

		for (File committeeFile : getCommitteeFiles(portfolioDirectory)) {
			if(buildTask.isCancelled()) {
				return;
			}
			try {
				CommitteeCreator.createCommitteeBudget(availableTabColors.peek(), p, committeeFile);
			} catch (Exception e) {
				buildTask.updateBuildMessage(e.toString());
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
		
		BudgetBuilder.writeTotals(overviewSheet);
		
		stylePortfolioOverview(overviewSheet);
		
		buildTask.updateBuildMessage("\t Overview done!");
		buildTask.updateBuildProgress(++currentProgress, totalProgress);
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
		
		CellStyle customPortfolioStyle = BudgetBuilder.getWorkbook().createCellStyle();
		customPortfolioStyle.cloneStyleFrom(Styles.PORTFOLIO_LABEL_STYLE);
		XSSFFont customPortfolioFont = BudgetBuilder.getWorkbook().createFont();
		customPortfolioFont.setFontHeightInPoints((short)10);
		customPortfolioFont.setFontName("Arial");
		customPortfolioFont.setColor(currentColor.getIndex());
		customPortfolioFont.setBold(false);
		customPortfolioFont.setItalic(false);
		
		customPortfolioStyle.setFont(customPortfolioFont);
		
		p.setPortfolioLabelStyle(customPortfolioStyle);
		
		for(int j = 1; j < overviewSheet.getLastRowNum()-1; j++) {
			XSSFRow dataRow = overviewSheet.getRow(j);
			
			dataRow.getCell(0).setCellStyle(customPortfolioStyle);
			
			dataRow.getCell(1).setCellStyle(Styles.COMMITTEE_LABEL_STYLE);
			
			for(int k = 2; k < dataRow.getLastCellNum(); k++)
				dataRow.getCell(k).setCellStyle(Styles.CURRENCY_CELL_STYLE);
		}
		
		//Total styles done in BudgetBuilder.writeTotals()
	}

	static File[] getCommitteeFiles(File portfolioDirectory) {
		return portfolioDirectory.listFiles(file -> {
							if (file.isHidden())
								return false;
							else if (file.getName().endsWith(".xlsx"))
								return true;
							return false;
						});
	}
}
