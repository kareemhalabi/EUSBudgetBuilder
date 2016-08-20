/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.controller;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import ca.mcgilleus.budgetbuilder.util.Styles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

import static ca.mcgilleus.budgetbuilder.controller.BudgetBuilder.*;
import static ca.mcgilleus.budgetbuilder.fxml.FileSelectController.getPreviousBudgetName;

public class PortfolioCreator {

	private static IndexedColors currentColor;

	/**
	 * Creates an EUS portfolio from a directory
	 * @param portfolioDirectory The root directory of the portfolio
	 * @param budget Budget to add portfolio to  
	 */
	public static void createPortfolio(File portfolioDirectory, EUSBudget budget) {

		Portfolio currentPortfolio = new Portfolio(portfolioDirectory.getName(), budget);

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

		buildTask.updateBuildProgress(++currentProgress, totalProgress);
	}

	public static void writeHeader(EUSBudget budget, XSSFSheet overviewSheet) {
		XSSFRow header = overviewSheet.createRow(0);

		XSSFCell portfolio = header.createCell(0, Cell.CELL_TYPE_STRING);
		portfolio.setCellValue("Portfolio");

		XSSFCell committee = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		committee.setCellValue("Committee/Function");

		XSSFCell rev = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		rev.setCellValue(budget.getBudgetYear() + " Revenues");

		XSSFCell exp = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		exp.setCellValue(budget.getBudgetYear() + " Expenses");

		XSSFCell budgetTitle = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		budgetTitle.setCellValue(budget.getBudgetYear() + " Budget");

		if(budget.hasPreviousYear()) {
			XSSFCell previousTitle = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
			previousTitle.setCellValue(getPreviousBudgetName());

			XSSFCell difference = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
			difference.setCellValue("Difference");
		}

		//Apply Header styles
		for(Cell cell : header) {
			cell.setCellStyle(Styles.HEADER_STYLE);
		}
	}

	public static void createMiscPortfolio(File miscPortfolioFile, EUSBudget budget) {

		Portfolio miscPortfolio = new Portfolio(miscPortfolioFile.getName().split(".xlsx")[0], budget);
		miscPortfolio.setMisc(true);

		buildTask.updateBuildMessage("Compiling misc portfolio: "+ miscPortfolio.getName());

		XSSFSheet destSheet = budget.getWb().createSheet(miscPortfolio.getName());
		writeHeader(budget, destSheet);

		//Prevents duplicate CellStyles for Portfolio overview and Budget overview
		if(miscPortfolio.getPortfolioLabelStyle() == null) {
			miscPortfolio.setPortfolioLabelStyle(Styles.getPortfolioLabelStyle(Styles.popTabColor()));
		}
		CellStyle customPortfolioStyle = miscPortfolio.getPortfolioLabelStyle();


		try(XSSFWorkbook miscWorkbook = new XSSFWorkbook(miscPortfolioFile)) {
			XSSFSheet srcSheet = miscWorkbook.getSheetAt(0);

			for(int i = 1; i <= srcSheet.getLastRowNum(); i++) {
				XSSFRow srcRow = srcSheet.getRow(i);
				XSSFRow destRow = destSheet.createRow(i);

				XSSFCell miscPortfolioLabel = destRow.createCell(PORTFOLIO_COL_INDEX, Cell.CELL_TYPE_STRING);
				miscPortfolioLabel.setCellValue(miscPortfolio.getName());
				miscPortfolioLabel.setCellStyle(customPortfolioStyle);

				XSSFCell functionName = destRow.createCell(COMMITTEE_COL_INDEX,Cell.CELL_TYPE_STRING);
				Cloner.cloneCell(srcRow.getCell(0), functionName, false);
				functionName.setCellStyle(Styles.COMMITTEE_LABEL_STYLE);

				XSSFCell functionRev = destRow.createCell(REV_COL_INDEX);
				Cloner.cloneCell(srcRow.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK),
						functionRev, false);
				functionRev.setCellStyle(Styles.CURRENCY_CELL_STYLE);

				XSSFCell functionExp = destRow.createCell(EXP_COL_INDEX);
				Cloner.cloneCell(srcRow.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK),
						functionExp, false);
				functionExp.setCellStyle(Styles.CURRENCY_CELL_STYLE);

				XSSFCell functionAmt = destRow.createCell(CURRENT_AMT_COL_INDEX, Cell.CELL_TYPE_FORMULA);
				functionAmt.setCellFormula(functionRev.getReference() + "+" + functionExp.getReference());
				functionAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

				CommitteeBudget committee = new CommitteeBudget(
						functionName.getStringCellValue(), //Function name
						getSheetCellReference(functionRev), //Revenues ref
						getSheetCellReference(functionExp), //Expenses ref
						miscPortfolio
				);

				if(budget.hasPreviousYear()){
					writePreviousCommittee(destRow, committee);
				}
			}

			if(budget.hasPreviousYear()) {
				writeInactiveCommittees(destSheet, miscPortfolio);
			}

			fixColumnWidths(destSheet);

		} catch (Exception e) {
			buildTask.updateBuildMessage(e.toString());
		}
		if(buildTask.isCancelled()) {
			return;
		}

		writeTotals(destSheet);
		buildTask.updateBuildProgress(++currentProgress, totalProgress);
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

			XSSFCell portfolioLabel = committeeRow.createCell(PORTFOLIO_COL_INDEX, Cell.CELL_TYPE_STRING);
			portfolioLabel.setCellValue(p.getName());
			portfolioLabel.setCellStyle(customPortfolioStyle);

			XSSFCell committeeName = committeeRow.createCell(COMMITTEE_COL_INDEX, Cell.CELL_TYPE_STRING);
			committeeName.setCellValue(committee.getName());
			committeeName.setCellStyle(Styles.COMMITTEE_LABEL_STYLE);

			XSSFCell committeeRev = committeeRow.createCell(REV_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			committeeRev.setCellFormula(committee.getRevRef());
			committeeRev.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell committeeExp = committeeRow.createCell(EXP_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			committeeExp.setCellFormula(committee.getExpRef());
			committeeExp.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell committeeAmt = committeeRow.createCell(CURRENT_AMT_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			committeeAmt.setCellFormula(committeeRev.getReference() + "+" + committeeExp.getReference());
			committeeAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			if(p.getEUSBudget().hasPreviousYear()) {
				writePreviousCommittee(committeeRow, committee);
			}
		}

		if(p.getEUSBudget().hasPreviousYear()) {
			writeInactiveCommittees(overviewSheet, p);
		}

		fixColumnWidths(overviewSheet);
	}

	public static void writePreviousCommittee(XSSFRow currentRow, CommitteeBudget currentCommittee) {
		EUSBudget budget = currentCommittee.getPortfolio().getEUSBudget();
		XSSFCell currentAmt = currentRow.getCell(4);

		XSSFCell previousAmt = currentRow.createCell(PREV_AMT_COL_INDEX, Cell.CELL_TYPE_NUMERIC);
		previousAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

		CommitteeBudget previousCommittee = budget.getPreviousCommittee(currentCommittee);
		if(previousCommittee != null) {
			previousAmt.setCellValue(previousCommittee.getPreviousAmt());

			Portfolio portfolioToRemove = previousCommittee.getPortfolio();

			//Remove committee so that only inactive ones are left
			portfolioToRemove.removeCommitteeBudget(previousCommittee);

			//Remove empty portfolios
			if(!portfolioToRemove.hasCommitteeBudgets()) {
				portfolioToRemove.getEUSBudget().removePortfolio(portfolioToRemove);
			}
		}

		XSSFCell difference = currentRow.createCell(DIFF_COL_INDEX, Cell.CELL_TYPE_FORMULA);
		difference.setCellFormula(currentAmt.getReference() + "-" + previousAmt.getReference());
		difference.setCellStyle(Styles.CURRENCY_CELL_STYLE);
	}

	private static void writeInactiveCommittees(XSSFSheet overviewSheet, Portfolio currentPortfolio) {
		Portfolio previousPortfolio = currentPortfolio.getEUSBudget().getPreviousPortfolio(currentPortfolio);
		if(previousPortfolio != null) {
			writeInactiveCommittees(overviewSheet, currentPortfolio, previousPortfolio);

			//Remove portfolio so that only inactive ones are left
			previousPortfolio.getEUSBudget().removePortfolio(previousPortfolio);
		}
	}

	static void writeInactiveCommittees(XSSFSheet overviewSheet, Portfolio currentPortfolio, Portfolio previousPortfolio) {
		for(CommitteeBudget inactiveCommittee : previousPortfolio.getCommitteeBudgets()) {
			XSSFRow committeeRow = overviewSheet.createRow(overviewSheet.getLastRowNum()+1);

			XSSFCell portfolioLabel = committeeRow.createCell(PORTFOLIO_COL_INDEX, Cell.CELL_TYPE_STRING);
			portfolioLabel.setCellValue(currentPortfolio.getName());
			if(currentPortfolio.getPortfolioLabelStyle() != null)
				portfolioLabel.setCellStyle(currentPortfolio.getPortfolioLabelStyle());

			XSSFCell committeeLabel = committeeRow.createCell(COMMITTEE_COL_INDEX, Cell.CELL_TYPE_STRING);
			committeeLabel.setCellValue(inactiveCommittee.getName());
			committeeLabel.setCellStyle(Styles.COMMITTEE_LABEL_STYLE);

			XSSFCell committeeRev = committeeRow.createCell(REV_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			committeeRev.setCellFormula("0");
			committeeRev.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell committeeExp = committeeRow.createCell(EXP_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			committeeExp.setCellFormula("0");
			committeeExp.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell currentAmt = committeeRow.createCell(CURRENT_AMT_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			currentAmt.setCellFormula(committeeRev.getReference() + "+" + committeeExp.getReference());
			currentAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell previousAmt = committeeRow.createCell(PREV_AMT_COL_INDEX, Cell.CELL_TYPE_NUMERIC);
			previousAmt.setCellValue(inactiveCommittee.getPreviousAmt());
			previousAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell difference = committeeRow.createCell(DIFF_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			difference.setCellFormula("-" + previousAmt.getReference());
			difference.setCellStyle(Styles.CURRENCY_CELL_STYLE);
		}
	}


	public static void fixColumnWidths(XSSFSheet overviewSheet) {
		XSSFRow header = overviewSheet.getRow(0);
		for(int i = 0; i < header.getLastCellNum(); i++) {
			overviewSheet.autoSizeColumn(i);

			//Extra space added to autoSizeColumn()
			int adjustedWidth = overviewSheet.getColumnWidth(i) + (256 * 3);

			//Prevents excessive column size
			if(adjustedWidth > 5200)
				adjustedWidth = 5200;

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
