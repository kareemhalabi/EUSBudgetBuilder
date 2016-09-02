package ca.mcgilleus.budgetbuilder.service;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.task.ValidationTask;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import ca.mcgilleus.budgetbuilder.util.Styles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;

import static ca.mcgilleus.budgetbuilder.service.BudgetBuilder.*;

/**
 * Service class for creating Portfolios and Misc Portfolios
 *
 * @author Kareem Halabi
 */
public class PortfolioCreator {

	private static IndexedColors currentColor;

	/**
	 * Creates a regular Portfolio from a directory
	 * <p>
	 * The Portfolio name is determined by the directory's name. Each portfolio is assigned a color for use on it's committee's
	 * tabs. In createPortfolioOverview() this color is used to create a custom XSSFCellStyle for the Portfolio label
	 * in overview sheets.
	 *
	 * @param portfolioDirectory The root directory of the portfolio
	 */
	public static void createPortfolio(File portfolioDirectory) {

		Portfolio currentPortfolio = new Portfolio(portfolioDirectory.getName(), budget);

		buildTask.updateBuildMessage("Compiling portfolio: " + currentPortfolio.getName());

		XSSFSheet overviewSheet = budget.getWb().createSheet(currentPortfolio.getName());

		currentColor = Styles.popTabColor();

		File[] committeeFiles = ValidationTask.getCommitteeFiles(portfolioDirectory);
		//TODO uncomment if CommitteeName is based off of file Name
//		Arrays.sort(committeeFiles);
		for (File committeeFile : committeeFiles) {
			try {
				CommitteeCreator.createCommitteeBudget(currentColor, currentPortfolio, committeeFile);
			} catch (Exception e) {
				buildTask.updateBuildMessage(e.toString());
			}
			if(buildTask.isCancelled()) {
				return;
			}
		}

		writeHeader(overviewSheet);
		createPortfolioOverview(currentPortfolio, overviewSheet);
		writeTotals(overviewSheet);

		buildTask.updateBuildProgress(++currentProgress, totalProgress);
	}

	/**
	 * Creates a Misc Portfolio and overview sheet from a top level xlsx file following the Misc Portfolio template
	 * <p>
	 * The Misc Portfolio name is determined by the xlsx filename. Misc Portfolios are given a color to create a custom
	 * XSSFCellStyle for the Portfolio label in overview sheets. Each row in the misc portfolio file is a Committee/Function
	 * which contains a cell for it's name, a cell for it's revenues and a cell for it's expenses
	 *
	 * @param miscPortfolioFile The top level xlsx file representing a Misc Portfolio
	 */
	public static void createMiscPortfolio(File miscPortfolioFile) {

		Portfolio miscPortfolio = new Portfolio(miscPortfolioFile.getName().split(".xlsx")[0], budget);
		miscPortfolio.setMisc(true);

		buildTask.updateBuildMessage("Compiling misc portfolio: "+ miscPortfolio.getName());

		XSSFSheet destSheet = budget.getWb().createSheet(miscPortfolio.getName());
		writeHeader(destSheet);

		//Prevents duplicate CellStyles for Portfolio overview and Budget overview
		if(miscPortfolio.getPortfolioLabelStyle() == null) {
			miscPortfolio.setPortfolioLabelStyle(Styles.getPortfolioLabelStyle(Styles.popTabColor()));
		}
		XSSFCellStyle customPortfolioStyle = miscPortfolio.getPortfolioLabelStyle();


		try(XSSFWorkbook miscWorkbook = new XSSFWorkbook(miscPortfolioFile)) {
			XSSFSheet srcSheet = miscWorkbook.getSheetAt(0);

			//   skip header
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
				functionAmt.setCellStyle(Styles.AMT_CURRENCY_CELL_STYLE);

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
	private static void createPortfolioOverview(Portfolio p, XSSFSheet overviewSheet) {

		//Prevents duplicate CellStyles for Portfolio overview and Budget overview
		if(p.getPortfolioLabelStyle() == null) {
			p.setPortfolioLabelStyle(Styles.getPortfolioLabelStyle(currentColor));
		}
		XSSFCellStyle customPortfolioStyle = p.getPortfolioLabelStyle();

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
			committeeAmt.setCellStyle(Styles.AMT_CURRENCY_CELL_STYLE);

			if(p.getEUSBudget().hasPreviousYear()) {
				writePreviousCommittee(committeeRow, committee);
			}
		}

		if(p.getEUSBudget().hasPreviousYear()) {
			writeInactiveCommittees(overviewSheet, p);
		}

		fixColumnWidths(overviewSheet);
	}

	/**
	 * Writes previous amount and difference cells for previous budget. Attempts to find the previous committee from the
	 * previous budget, if it exists, the previous amount and difference is written to the currentRow. The previous committee is then
	 * removed from previous budget so that after all previous committees have been matched with current versions, the
	 * inactive committees are the ones left in the previous budget
	 * @param currentRow the row to write the previous amt and difference
	 * @param currentCommittee the committee to match to a previous version
	 */
	private static void writePreviousCommittee(XSSFRow currentRow, CommitteeBudget currentCommittee) {
		EUSBudget budget = currentCommittee.getPortfolio().getEUSBudget();

		XSSFCell currentAmt = currentRow.getCell(CURRENT_AMT_COL_INDEX);

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

	/**
	 * Writes inactive committee rows for portfolio overview sheets. Attempts to find the previous portfolio and then
	 * uses overloaded version of this method to write the appropriate cells. The portfolio is then removed
	 * it if found so that after all previous portfolios have been matched, inactive ones are left.
	 * @param overviewSheet the overview sheet to write inactive committees to.
	 * @param currentPortfolio the portfolio to match to a previous verision
	 */
	private static void writeInactiveCommittees(XSSFSheet overviewSheet, Portfolio currentPortfolio) {
		Portfolio previousPortfolio = currentPortfolio.getEUSBudget().getPreviousPortfolio(currentPortfolio);
		if(previousPortfolio != null) {
			writeInactiveCommittees(overviewSheet, currentPortfolio, previousPortfolio);

			//Remove portfolio so that only inactive ones are left
			previousPortfolio.getEUSBudget().removePortfolio(previousPortfolio);
		}
	}

	/**
	 * Writes the rows for inactive committees in a portfolio. If the portfolio is active it's previously set cell style
	 * is used for labeling, otherwise the generic label is applied. Current revenues and expesnses are set to 0 and
	 * budget net formula is still written in case workbook is modified mannually after being generated. Previous amounts
	 * and difference formula are also set
	 * @param overviewSheet the overview sheet to write inactive committees to
	 * @param currentPortfolio the current version of the portfolio
	 * @param previousPortfolio the previous version of the portfolio
	 */
	static void writeInactiveCommittees(XSSFSheet overviewSheet, Portfolio currentPortfolio, Portfolio previousPortfolio) {
		for(CommitteeBudget inactiveCommittee : previousPortfolio.getCommitteeBudgets()) {
			XSSFRow committeeRow = overviewSheet.createRow(overviewSheet.getLastRowNum()+1);

			XSSFCell portfolioLabel = committeeRow.createCell(PORTFOLIO_COL_INDEX, Cell.CELL_TYPE_STRING);
			portfolioLabel.setCellValue(currentPortfolio.getName());
			if(currentPortfolio.getPortfolioLabelStyle() != null)
				portfolioLabel.setCellStyle(currentPortfolio.getPortfolioLabelStyle());
			else
				portfolioLabel.setCellStyle(Styles.PORTFOLIO_LABEL_STYLE);

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
			currentAmt.setCellStyle(Styles.AMT_CURRENCY_CELL_STYLE);

			XSSFCell previousAmt = committeeRow.createCell(PREV_AMT_COL_INDEX, Cell.CELL_TYPE_NUMERIC);
			previousAmt.setCellValue(inactiveCommittee.getPreviousAmt());
			previousAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

			XSSFCell difference = committeeRow.createCell(DIFF_COL_INDEX, Cell.CELL_TYPE_FORMULA);
			difference.setCellFormula("-" + previousAmt.getReference());
			difference.setCellStyle(Styles.CURRENCY_CELL_STYLE);
		}
	}
}
