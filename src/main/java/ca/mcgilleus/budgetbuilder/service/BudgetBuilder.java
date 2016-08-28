/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.service;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.task.BuildTask;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import ca.mcgilleus.budgetbuilder.util.Styles;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Calendar;
import java.util.Date;

import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.previousBudgetFile;
import static ca.mcgilleus.budgetbuilder.service.PortfolioCreator.*;


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

	//TODO add comments here
	public static void rebuildPreviousBudget() {

		buildTask.updateBuildMessage("Rebuilding previous budget");

		EUSBudget previousBudget = createBudget();
		budget.setPreviousYear(previousBudget);

		try (XSSFWorkbook previousWorkbook = new XSSFWorkbook(previousBudgetFile)) {
			XSSFSheet previousBudgetSheet = previousWorkbook.getSheetAt(0);

			String previousRowPortfolioLabel = "";
			Portfolio previousPortfolio = null;
			//					Ignore total rows
			for(int i = 1; i <= previousBudgetSheet.getLastRowNum()-2; i++) {

				XSSFRow committeeRow = previousBudgetSheet.getRow(i);

				//Skip a committee if it has no expense or revenue activity
				if(committeeRow.getCell(2).getNumericCellValue() == 0 &&
						committeeRow.getCell(3).getNumericCellValue() == 0) {
					continue;
				}

				String portfolioLabel = committeeRow.getCell(0).getStringCellValue();

				if(!portfolioLabel.trim().toUpperCase().equals(previousRowPortfolioLabel.trim().toUpperCase())) {
					previousPortfolio = new Portfolio(portfolioLabel, previousBudget);
					previousRowPortfolioLabel = portfolioLabel;
				}

				String committeeLabel = committeeRow.getCell(1).getStringCellValue();
				double committeeAmt = committeeRow.getCell(4).getNumericCellValue();

				assert previousPortfolio != null;
				new CommitteeBudget(committeeLabel, "", "", previousPortfolio).setPreviousAmt(committeeAmt);

			}
		} catch(Exception e) {
			buildTask.updateBuildMessage(e.toString());
		}
	}

	public static void createBudgetOverview(XSSFSheet budgetOverview) {
		buildTask.updateBuildMessage("Compiling EUS Budget Overview");
		buildTask.updateBuildProgress(++currentProgress, totalProgress);

		writeHeader(budget, budgetOverview);

		budget.getWb().setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
		int destIndex = 1;
		for(Portfolio p : budget.getPortfolios()) {
			XSSFSheet portfolioOverview = budget.getWb().getSheet(p.getName());
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

	static String getSheetCellReference(XSSFCell cell) {
		return "\'" + cell.getSheet().getSheetName() + "\'!" + cell.getReference();
	}
}