package ca.mcgilleus.budgetbuilder.service;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.task.BuildTask;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import ca.mcgilleus.budgetbuilder.util.Styles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Calendar;
import java.util.Date;

import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.previousBudgetFile;
import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.previousBudgetName;

/**
 * This class is the main service that performs the budget building.
 */
public class BudgetBuilder {

	static final int PORTFOLIO_COL_INDEX = 0;
	static final int COMMITTEE_COL_INDEX = 1;
	static final int REV_COL_INDEX = 2;
	static final int EXP_COL_INDEX = 3;
	static final int CURRENT_AMT_COL_INDEX = 4;
	static final int PREV_AMT_COL_INDEX = 5;
	static final int DIFF_COL_INDEX = 6;

	public static BuildTask buildTask;
	public static int totalCommitteeFiles;
	public static int totalMiscPortfolios;
	public static int totalProgress;
	public static int currentProgress;

	public static EUSBudget budget;

	/**
	 * Creates a new budget based on calendar year of Fall Semester
	 * @return a new budget based on calendar year of Fall Semester
	 */
	public static EUSBudget createBudget() {
		Calendar c = Calendar.getInstance();
		Date budgetYear;
		if(c.get(Calendar.MONTH) <= Calendar.FEBRUARY)
			c.add(Calendar.YEAR, -1);

		budgetYear = c.getTime();

		return new EUSBudget(budgetYear);
	}

	/**
	 * Reconstructs a model of the previous year's budget from it's excel file.
	 * <p>
	 * A new EUSBudget object is created and set as this year's previous. Then cycling through each row starting after
	 * the header (Row 2) until the last row (2 before the Total Row) the portfolios, committees and budget net cells
	 * are read. Any committees having 0 expenses or revenues are deemed inactive for two consecutive budgets and are skipped.
	 * <p>
	 * On any given row, if the portfolio label does not match the previous row's portfolio label, a new Portfolio is
	 * created and the Committee on this row is added to the new Portfolio. If the portfolio label does match the previous,
	 * the Committee is added to the pre-existing Portfolio. (This is why I have the fields previousRowPortfolioLabel and
	 * previousPortfolio)
	 */
	public static void rebuildPreviousBudget() {

		buildTask.updateBuildMessage("Rebuilding previous budget");

		EUSBudget previousBudget = createBudget();
		budget.setPreviousYear(previousBudget);

		try (XSSFWorkbook previousWorkbook = new XSSFWorkbook(previousBudgetFile)) {
			XSSFSheet previousBudgetSheet = previousWorkbook.getSheetAt(0);

			String previousRowPortfolioLabel = "";
			Portfolio previousPortfolio = null;
			//	skip header		Ignore total rows
			for(int i = 1; i <= previousBudgetSheet.getLastRowNum()-2; i++) {

				XSSFRow committeeRow = previousBudgetSheet.getRow(i);

				//Skip a committee if it has no expense or revenue activity
				if(committeeRow.getCell(REV_COL_INDEX).getNumericCellValue() == 0 &&
						committeeRow.getCell(EXP_COL_INDEX).getNumericCellValue() == 0) {
					continue;
				}

				String portfolioLabel = committeeRow.getCell(PORTFOLIO_COL_INDEX).getStringCellValue();

				if(!portfolioLabel.trim().toUpperCase().equals(previousRowPortfolioLabel.trim().toUpperCase())) {
					previousPortfolio = new Portfolio(portfolioLabel, previousBudget);
					previousRowPortfolioLabel = portfolioLabel;
				}

				String committeeLabel = committeeRow.getCell(COMMITTEE_COL_INDEX).getStringCellValue();
				double committeeAmt = committeeRow.getCell(CURRENT_AMT_COL_INDEX).getNumericCellValue();

				assert previousPortfolio != null;
				new CommitteeBudget(committeeLabel, "", "", previousPortfolio).setPreviousAmt(committeeAmt);

			}
		} catch(Exception e) {
			buildTask.updateBuildMessage(e.toString());
		}
	}

	/**
	 * Creates the master budget overview sheet.
	 * <p>
	 *  After writing the headers, each portfolio overview is cycled through copying the relevant information.
	 *  Portfolio and Committee labels are Cloned. References to revenues and expenses from Misc Portfolio Overviews
	 *  are directly copied over whereas for regular Portfolios, the revenue and expense references are copied from
	 *  the Committees themselves. Budget net cell formulas are written in-place (i.e not a copied reference)
	 * <p>
	 *  If a previous budget exists, the previous amount references are copied over and the difference cell formulas
	 *  are written in-place (i.e not a copied reference). Any inactive portfolios are written last.
	 * <p>
	 *  Finally totals are written, columns are auto-sized and a free pane is created for the first row
	 *
	 * @param budgetOverview the sheet to write the overview to.
	 */
	public static void createBudgetOverview(XSSFSheet budgetOverview) {
		buildTask.updateBuildMessage("Compiling EUS Budget Overview");
		buildTask.updateBuildProgress(++currentProgress, totalProgress);

		writeHeader(budgetOverview);

		budget.getWb().setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

		//Keeps track of current row in budgetOverview. Initialized to 1 because header is skipped.
		int destIndex = 1;
		for(Portfolio p : budget.getPortfolios()) {
			XSSFSheet portfolioOverview = budget.getWb().getSheet(p.getName());
			//	   skip header		Ignore total rows
			for(int srcIndex = 1; srcIndex <= portfolioOverview.getLastRowNum()-2; srcIndex++) {

				XSSFRow srcRow = portfolioOverview.getRow(srcIndex);
				XSSFRow destRow = budgetOverview.createRow(destIndex);

				Cloner.cloneCell(srcRow.getCell(PORTFOLIO_COL_INDEX),destRow.createCell(PORTFOLIO_COL_INDEX), true);
				Cloner.cloneCell(srcRow.getCell(COMMITTEE_COL_INDEX),destRow.createCell(COMMITTEE_COL_INDEX), true);

				XSSFCell revCell = destRow.createCell(REV_COL_INDEX, Cell.CELL_TYPE_FORMULA);
				XSSFCell expCell = destRow.createCell(EXP_COL_INDEX, Cell.CELL_TYPE_FORMULA);
				if(p.isMisc()){
					revCell.setCellFormula(getSheetCellReference(srcRow.getCell(REV_COL_INDEX)));
					expCell.setCellFormula(getSheetCellReference(srcRow.getCell(EXP_COL_INDEX)));
				} else {
					revCell.setCellFormula(srcRow.getCell(REV_COL_INDEX).getCellFormula());
					expCell.setCellFormula(srcRow.getCell(EXP_COL_INDEX).getCellFormula());
				}
				revCell.setCellStyle(Styles.CURRENCY_CELL_STYLE);
				expCell.setCellStyle(Styles.CURRENCY_CELL_STYLE);

				XSSFCell currentAmt = destRow.createCell(CURRENT_AMT_COL_INDEX, Cell.CELL_TYPE_FORMULA);
				currentAmt.setCellFormula(revCell.getReference() + "+" + expCell.getReference());
				currentAmt.setCellStyle(Styles.AMT_CURRENCY_CELL_STYLE);

				if(budget.hasPreviousYear()) {
					XSSFCell previousAmt = destRow.createCell(PREV_AMT_COL_INDEX, Cell.CELL_TYPE_FORMULA);
					previousAmt.setCellFormula(getSheetCellReference(srcRow.getCell(PREV_AMT_COL_INDEX)));
					previousAmt.setCellStyle(Styles.CURRENCY_CELL_STYLE);

					XSSFCell difference = destRow.createCell(DIFF_COL_INDEX, Cell.CELL_TYPE_FORMULA);
					difference.setCellFormula(currentAmt.getReference() + "-" + previousAmt.getReference());
					difference.setCellStyle(Styles.CURRENCY_CELL_STYLE);
				}
				destIndex++;
			}
		}

		//Add any inactive portfolios if they exist
		if(budget.hasPreviousYear()) {
			for(Portfolio inactivePortfolio : budget.getPreviousYear().getPortfolios()) {
				PortfolioCreator.writeInactiveCommittees(budgetOverview, inactivePortfolio, inactivePortfolio);
			}
		}

		writeTotals(budgetOverview);

		fixColumnWidths(budgetOverview);

		budgetOverview.createFreezePane(0,1);
	}

	/**
	 * Generates a sheet-based reference to a cell
	 * @param cell the cell to reference to
	 * @return a String representing the sheet-based reference to passed-in cell
	 */
	static String getSheetCellReference(XSSFCell cell) {
		return "\'" + cell.getSheet().getSheetName() + "\'!" + cell.getReference();
	}

	/**
	 * Creates a header for an overviewSheet
	 * Headings are as follows:
	 * Portfolio	Committee/Function	   Revenues		Expenses	Current Budget	 Previous Budget	Difference
	 * 																				  (if exists)       (if exists)
	 * @param overviewSheet the sheet to write the header to
	 */
	static void writeHeader(XSSFSheet overviewSheet) {
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
			previousTitle.setCellValue(previousBudgetName);

			XSSFCell difference = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
			difference.setCellValue("Difference");
		}

		//Apply Header styles
		for(Cell cell : header) {
			cell.setCellStyle(Styles.HEADER_STYLE);
		}
	}

	/**
	 * Adjusts the width of all columns in a sheet such that cell contents don't overlap into their neighbours.
	 * The standard autosize method is too tight so some space (arbitrarily set at 256*3) is added. If the autosized
	 * column width exceeds an arbitrary size of 5200, it is capped at 5200
	 *
	 * @param overviewSheet the sheet whose columns will be adjusted
	 */
	static void fixColumnWidths(XSSFSheet overviewSheet) {
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
	 * Creates ans styles total cells representing the sum of all numerical cells above it.
	 * There is a blank row between the last entry and total row.
	 *
	 * @param overviewSheet the sheet to write totals to
	 */
	static void writeTotals(XSSFSheet overviewSheet) {

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
}