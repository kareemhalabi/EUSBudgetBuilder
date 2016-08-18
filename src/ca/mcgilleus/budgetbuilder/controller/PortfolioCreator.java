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
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;

import static ca.mcgilleus.budgetbuilder.controller.BudgetBuilder.*;

public class PortfolioCreator {

	private static IndexedColors currentColor;
	private static Portfolio currentPortfolio;

	/**
	 * Creates an EUS portfolio from a directory
	 * @param portfolioDirectory The root directory of the portfolio
	 * @param budget Budget to add portfolio to  
	 */
	public static void createPortfolio(File portfolioDirectory, EUSBudget budget) {

		currentPortfolio = new Portfolio(portfolioDirectory.getName(), budget);

		buildTask.updateBuildMessage("Compiling portfolio: " + currentPortfolio.getName());

		XSSFSheet overviewSheet = budget.getWb().createSheet(currentPortfolio.getName());

		currentColor = Styles.popTabColor();

		for (File committeeFile : getCommitteeFiles(portfolioDirectory)) {
			try {
				CommitteeCreator.createCommitteeBudget(currentColor, currentPortfolio, committeeFile);
			} catch (Exception e) {
				buildTask.updateBuildMessage(e.toString());
			}
			if(buildTask.isCancelled()) {
				return;
			}
		}

		writeHeader(budget, overviewSheet);
		createPortfolioOverview(currentPortfolio, overviewSheet);
		writeTotals(overviewSheet);

		buildTask.updateBuildMessage("\t Overview done!");
		buildTask.updateBuildProgress(++currentProgress, totalProgress);
	}

	public static void writeHeader(EUSBudget budget, XSSFSheet overviewSheet) {
		XSSFRow header = overviewSheet.createRow(0);

		XSSFCell portfolio = header.createCell(0, Cell.CELL_TYPE_STRING);
		portfolio.setCellValue("Portfolio");

		XSSFCell committee = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		committee.setCellValue("Committee");

		XSSFCell rev = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		rev.setCellValue(budget.getBudgetYear() + " Revenues");

		XSSFCell exp = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		exp.setCellValue(budget.getBudgetYear() + " Expenses");

		XSSFCell budgetTitle = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		budgetTitle.setCellValue(budget.getBudgetYear() + " Budget");

		//Apply Header styles
		for(Cell cell : header) {
			cell.setCellStyle(Styles.HEADER_STYLE);
		}
	}

	/**
	 * Creates and styles overview rows outlining each Portfolio, Committee and the amount they requested.
	 * The Portfolio and Committee name cells are string values whereas the Committee amounts are
	 * cell references to the amount requested on the Committee's sheet.
	 *
	 * @param p The porfirio overview to create
	 * @param overviewSheet The sheet to write the overview rows
	 */
	public static void createPortfolioOverview(Portfolio p, XSSFSheet overviewSheet) {

		//Prevents duplicate CellStyles for Portfolio overview and Budget overview
		if(p.getPortfolioLabelStyle() == null) {
			p.setPortfolioLabelStyle(Styles.getPortfolioLabelStyle(currentColor));
		}
		CellStyle customPortfolioStyle = p.getPortfolioLabelStyle();

		for(CommitteeBudget committee : p.getCommitteeBudgets()) {
			//getLastRowNum()+1 because starts on second row
			XSSFRow committeeRow = overviewSheet.createRow(overviewSheet.getLastRowNum()+1);

			XSSFCell portfolioLabel = committeeRow.createCell(0, Cell.CELL_TYPE_STRING);
			portfolioLabel.setCellValue(p.getName());
			portfolioLabel.setCellStyle(customPortfolioStyle);

			XSSFCell committeeName = committeeRow.createCell(committeeRow.getLastCellNum(), Cell.CELL_TYPE_STRING);
			committeeName.setCellValue(committee.getName());
			committeeName.setCellStyle(Styles.COMMITTEE_LABEL_STYLE);

			XSSFCell committeeRev = committeeRow.createCell(committeeRow.getLastCellNum(), Cell.CELL_TYPE_FORMULA);
			committeeRev.setCellFormula(committee.getRevRef());
			committeeRev.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell committeeExp = committeeRow.createCell(committeeRow.getLastCellNum(), Cell.CELL_TYPE_FORMULA);
			committeeExp.setCellFormula(committee.getExpRef());
			committeeExp.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell committeeAmt = committeeRow.createCell(committeeRow.getLastCellNum(), Cell.CELL_TYPE_FORMULA);
			committeeAmt.setCellFormula(committee.getRevRef() + "+" + committee.getExpRef());
			committeeAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);
		}

		XSSFRow header = overviewSheet.getRow(0);
		//Fix Column Widths
		for(int i = 0; i < header.getLastCellNum(); i++) {
			overviewSheet.autoSizeColumn(i);
			int adjustedWidth = overviewSheet.getColumnWidth(i) + 256*3;

			if(adjustedWidth > 5000)
				adjustedWidth = 5000;

			overviewSheet.setColumnWidth(i, adjustedWidth);
		}
	}

	/**
	 * Creates total cells representing the sum of all numerical cells above it.
	 * There is a blank row between the last entry and total row.
	 * Total cell styling is also done in this method
	 *
	 * @param overviewSheet the sheet to write totals to
	 */
	public static void writeTotals(XSSFSheet overviewSheet) {

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

		//Apply Total styles
		totalLabel.setCellStyle(Styles.TOTAL_LABEL_STYLE);

		for(int l = 2; l < totals.getLastCellNum(); l++)
			totals.getCell(l).setCellStyle(Styles.TOTAL_CELL_STYLE);
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
