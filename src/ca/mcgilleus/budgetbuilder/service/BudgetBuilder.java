/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.service;

import ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController;
import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import ca.mcgilleus.budgetbuilder.util.Styles;
import javafx.concurrent.Task;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.getPreviousBudgetFile;
import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.getSelectedDirectory;
import static ca.mcgilleus.budgetbuilder.service.PortfolioCreator.*;
import static ca.mcgilleus.budgetbuilder.util.Styles.initStyles;


public class BudgetBuilder {

	public static final int PORTFOLIO_COL_INDEX = 0;
	public static final int COMMITTEE_COL_INDEX = 1;
	public static final int REV_COL_INDEX = 2;
	public static final int EXP_COL_INDEX = 3;
	public static final int CURRENT_AMT_COL_INDEX = 4;
	public static final int PREV_AMT_COL_INDEX = 5;
	public static final int DIFF_COL_INDEX = 6;

	public static BuildTask buildTask;
	private static int totalCommitteeFiles;
	private static int totalMiscPortfolios;
	static int totalProgress;
	static int currentProgress;

	private static EUSBudget budget;

	public static Task getValidationTask() {

		return new Task() {

			private String errors = "";

			@Override
			protected Boolean call() throws Exception {

				final List<File> miscPortfoliosToCheck = getMiscPortfolioFilesToCheck(getSelectedDirectory());
				final List<File> committeeFilesToCheck = getCommitteeFilesToCheck(getSelectedDirectory());

				totalMiscPortfolios = miscPortfoliosToCheck.size();
				totalCommitteeFiles = committeeFilesToCheck.size();

				//committeeFilesToCheck may be null due to Task cancellation or some other error
				if (committeeFilesToCheck == null)
					return false;

				//																+ 1 for output file
				double totalProgress = totalMiscPortfolios + totalCommitteeFiles + 1;
				if(getPreviousBudgetFile() != null) {
					totalProgress++; // + 1 for previous file
				}
				double currentProgress = 0;

				//Check misc portfolios for errors
				for (File f : miscPortfoliosToCheck) {

					try (XSSFWorkbook workbook = new XSSFWorkbook(f)) {

						XSSFSheet miscSheet = workbook.getSheetAt(0);

						//Check for positive revenues and negative expenses
						for (int i = 1; i <= miscSheet.getLastRowNum(); i++) {

							XSSFRow row = miscSheet.getRow(i);

							XSSFCell rev = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							if(rev.getNumericCellValue() < 0) {
								CellReference revRef = new CellReference(rev);
								//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
								errors += "- Revenue cell " + revRef.getCellRefParts()[2] + revRef.getCellRefParts()[1] + " in \""
										 + f.getName() + "\" is not positive\n";
							}

							XSSFCell exp = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							if(exp.getNumericCellValue() > 0) {
								CellReference expRef = new CellReference(exp);
								//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
								errors += "- Expense cell " + expRef.getCellRefParts()[2] + expRef.getCellRefParts()[1] + " in \""
										+ f.getName() + "\" is not negative\n";
							}
						}
					} catch (Exception e) {
						if (e.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
							errors += "- Close file: \"" + f.getName() + "\"\n";
						}
						else {
							errors += e.toString();
						}
					}
					if (isCancelled()) {
						updateMessage("Cancelled");
						return false;
					}
					updateProgress(++currentProgress, totalProgress);
				}

				//Check portfolios for errors
				for (File f : committeeFilesToCheck) {

					//Check if file is already open
					try (XSSFWorkbook workbook = new XSSFWorkbook(f)) {

						//Check names
						if (workbook.getName("REV") == null) {
							errors += "- REV cell name missing in \"" + f.getParentFile().getName() + "\\" + f.getName() + "\"\n";
						} else {
							//Check if Revenues are positive
							CellReference revRef = new CellReference(workbook.getName("REV").getRefersToFormula());
							assert revRef.getSheetName() != null;
							XSSFRow revRow = workbook.getSheet(revRef.getSheetName()).getRow(revRef.getRow());
							XSSFCell revCell = revRow.getCell(revRef.getCol());
							if(revCell.getNumericCellValue() < 0) {
								//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
								errors += "- REV cell " + revRef.getCellRefParts()[2] + revRef.getCellRefParts()[1] + " in \""
										+ f.getParentFile().getName() + "\\" + f.getName() + "\" is not positive\n";
							}
						}
						if (workbook.getName("EXP") == null) {
							errors += "- EXP cell name missing in \"" + f.getName() + "\"\n";
						} else {
							//Check if Expenses are negative
							CellReference expRef = new CellReference(workbook.getName("EXP").getRefersToFormula());
							assert expRef.getSheetName() != null;
							XSSFRow expRow = workbook.getSheet(expRef.getSheetName()).getRow(expRef.getRow());
							XSSFCell expCell = expRow.getCell(expRef.getCol());
							if(expCell.getNumericCellValue() > 0) {
								//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
								errors += "- EXP cell " + expRef.getCellRefParts()[2] + expRef.getCellRefParts()[1] + " in \""
										+ f.getParentFile().getName() + "\\" + f.getName() + "\" is not negative\n";
							}
						}

						if (workbook.getName("NAME") == null) {
							errors += "- NAME cell name missing in \"" + f.getName() + "\"\n";
						}

					} catch (Exception e) {
						if (e.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
							errors += "- Close file: \"" + f.getParentFile().getName() + "\\" + f.getName() + "\"\n";
						}
						else {
							errors += e.toString();
						}
					}
					if (isCancelled()) {
						updateMessage("Cancelled");
						return false;
					}
					updateProgress(++currentProgress, totalProgress);
				}

				//Check if outputFile is open
				File outputFile = FileSelectController.getOutputFile();
				if(outputFile.exists()) {
					//Don't understand why this is necessary, prevents a Zip bomb error
					ZipSecureFile.setMinInflateRatio(0.001);

					//noinspection EmptyTryBlock
					try (XSSFWorkbook workbook = new XSSFWorkbook(outputFile)) {
					} catch (Exception e) {
						if (e.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
							errors += "- Close file: \"" + outputFile.getName() + "\"\n";
						}
						else {
							errors += e.toString();
						}
					}
					if (isCancelled()) {
						updateMessage("Cancelled");
						return false;
					}
				}

				//Check if Previous Budget is open
				if(getPreviousBudgetFile() != null) {
					try (XSSFWorkbook workbook = new XSSFWorkbook(getPreviousBudgetFile())) {
					} catch (Exception e) {
						if (e.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
							errors += "- Close file: \"" + outputFile.getName() + "\"\n";
						}
						else {
							errors += e.toString();
						}
					}
					if (isCancelled()) {
						updateMessage("Cancelled");
						return false;
					}

				}

				if(errors.trim().length() != 0) {
					updateMessage("Errors occurred:\n" + errors);
					return false;
				}
				return true;
			}

			private List<File> getMiscPortfolioFilesToCheck(File rootDirectory) {
				return Arrays.asList(getCommitteeFiles(rootDirectory));
			}

			private List<File> getCommitteeFilesToCheck(File rootDirectory) {

				ArrayList<File> filesToCheck = new ArrayList<>();

				//Get Portfolio directories ignoring files and hidden directories
				File[] portfolioDirectories = getPortfolioDirectories(rootDirectory);

				assert portfolioDirectories != null;
				for(File portfolioDirectory : portfolioDirectories) {

					if(isCancelled()) {
						updateMessage("Cancelled");
						return null;
					}

					//Get Committee excel committeeFiles ignoring any hidden committeeFiles
					//noinspection ConstantConditions
					File[] committeeFiles = PortfolioCreator.getCommitteeFiles(portfolioDirectory);

					assert committeeFiles != null;
					if(committeeFiles.length == 0)
						errors += "- Portfolio \"" + portfolioDirectory.getName() + "\" contains no committee budgets\n";
					Collections.addAll(filesToCheck, committeeFiles);
				}
				return filesToCheck;
			}
		};
	}

	public static class BuildTask extends Task{

		@Override
		protected Boolean call() throws Exception {

			buildTask = this;

			budget = createBudget();

			initStyles();

			if(getPreviousBudgetFile() != null)
				rebuildPreviousBudget();
			if(isCancelled()) {
				updateBuildMessage("Cancelled");
				return false;
			}

			XSSFSheet budgetOverview = budget.getWb().createSheet(budget.getBudgetYear() + " Budget");

			File[] miscPortfolioFiles = getCommitteeFiles(getSelectedDirectory());
			Arrays.sort(miscPortfolioFiles);

			File[] portfolioDirectories = getPortfolioDirectories(getSelectedDirectory());
			Arrays.sort(portfolioDirectories);

			//Total progress: total misc portfolios + total committee sheets + each portfolio overview + budget overview
			totalProgress = totalMiscPortfolios + totalCommitteeFiles + portfolioDirectories.length + 1;
			currentProgress = 0;

			//Cycle through each misc portfolio
			for (File miscPorfolio : miscPortfolioFiles) {
				PortfolioCreator.createMiscPortfolio(miscPorfolio, budget);
				if(isCancelled()) {
					updateBuildMessage("Cancelled");
					return false;
				}
			}

			//Cycle through each portfolio folder
			for(File portfolioDirectory : portfolioDirectories) {
				PortfolioCreator.createPortfolio(portfolioDirectory, budget);
				if(isCancelled()) {
					updateBuildMessage("Cancelled");
					return false;
				}
			}

			createBudgetOverview(budgetOverview);

			FileOutputStream fileOut;
			try {
				fileOut = new FileOutputStream(FileSelectController.getOutputFile());
				budget.getWb().write(fileOut);
				fileOut.close();
				budget.getWb().close();
				updateBuildMessage("Done! Press finish to open");
			} catch (IOException e) {
				updateBuildMessage(e.toString());
			}

			return true;
		}

		public void updateBuildMessage(String message) {
			updateMessage(message);
		}

		public void updateBuildProgress(double workDone, double max) {
			updateProgress(workDone, max);
		}
	}

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

	public static XSSFWorkbook getWorkbook() {
		return budget.getWb();
	}

	/**
	 * Returns all File objects that qualify as Portfolio directories.
	 * A File qualifies as a Portfolio if it is not hidden and is a directory
	 *
	 * @param rootDirectory the root directory containing portfolio subdirectories
	 * @return a File array containing portfolio directories
	 */
	private static File[] getPortfolioDirectories(File rootDirectory) {
		return rootDirectory.listFiles(file -> {
			if(file.isHidden())
				return false;
			else if(file.isDirectory())
				return true;
			return false;
		});
	}

	//TODO add comments here
	private static void rebuildPreviousBudget() {

		buildTask.updateBuildMessage("Rebuilding previous budget");

		EUSBudget previousBudget = createBudget();
		budget.setPreviousYear(previousBudget);

		try (XSSFWorkbook previousWorkbook = new XSSFWorkbook(getPreviousBudgetFile())) {
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

	private static void createBudgetOverview(XSSFSheet budgetOverview) {
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

	public static String getSheetCellReference(XSSFCell cell) {
		return "\'" + cell.getSheet().getSheetName() + "\'!" + cell.getReference();
	}
}