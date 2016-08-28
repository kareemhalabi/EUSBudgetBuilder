package ca.mcgilleus.budgetbuilder.task;

import ca.mcgilleus.budgetbuilder.service.BudgetBuilder;
import javafx.concurrent.Task;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.*;

/**
 * This Task is responsible for ensuring all files are valid for processing. Checks that all MiscPortfolio files,
 * Committee files, the output file(if exists) and the previous budget file (if exists and selected) are closed.
 *
 * In addition, empty Portfolio directories are checked, MiscPortfolio are checked for positive revenue cells and
 * negative revenue cells and Committee files are checked for containing REV, EXP and NAME names and that revenues are
 * positive, expenses are negative.
 */
public class ValidationTask extends Task {

	private String errors = "";

	@Override
	protected Boolean call() throws Exception {

		//Both are lists for consistency
		final List<File> miscPortfoliosToCheck = Arrays.asList(getCommitteeFiles(selectedDirectory));
		final List<File> committeeFilesToCheck = getCommitteeFilesToCheck(selectedDirectory);

		//committeeFilesToCheck may be null due to Task cancellation or some other error
		if (committeeFilesToCheck == null) {
			return false;
		}

		BudgetBuilder.totalMiscPortfolios = miscPortfoliosToCheck.size();
		BudgetBuilder.totalCommitteeFiles = committeeFilesToCheck.size();

		//																			+ 1 for output file
		double totalProgress = BudgetBuilder.totalMiscPortfolios + BudgetBuilder.totalCommitteeFiles + 1;
		if(previousBudgetFile != null) {
			totalProgress++; // + 1 for previous file
		}
		double currentProgress = 0;

		//Check misc portfolios for errors
		for (File miscFile : miscPortfoliosToCheck) {

			try (XSSFWorkbook workbook = new XSSFWorkbook(miscFile)) {

				XSSFSheet miscSheet = workbook.getSheetAt(0);

				for (int i = 1; i <= miscSheet.getLastRowNum(); i++) {

					XSSFRow row = miscSheet.getRow(i);

					//Check for positive revenues
					XSSFCell rev = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if(rev.getNumericCellValue() < 0) {
						CellReference revRef = new CellReference(rev);
						//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
						errors += "- Revenue cell " + revRef.getCellRefParts()[2] + revRef.getCellRefParts()[1] + " in \""
								+ miscFile.getName() + "\" is not positive\n";
					}

					//Check for negative expenses
					XSSFCell exp = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if(exp.getNumericCellValue() > 0) {
						CellReference expRef = new CellReference(exp);
						//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
						errors += "- Expense cell " + expRef.getCellRefParts()[2] + expRef.getCellRefParts()[1] + " in \""
								+ miscFile.getName() + "\" is not negative\n";
					}
				}
			} catch (Exception e) {
				//Check if file is open
				if (e.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
					errors += "- Close file: \"" + miscFile.getName() + "\"\n";
				}
				else {
					//Log all other errors
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
		for (File committeeFile : committeeFilesToCheck) {

			try (XSSFWorkbook workbook = new XSSFWorkbook(committeeFile)) {

				//Check REV Name
				if (workbook.getName("REV") == null) {
					errors += "- REV cell name missing in \"" + committeeFile.getParentFile().getName() + "\\" + committeeFile.getName() + "\"\n";
				}
				else {
					//Check if Revenues are positive
					CellReference revRef = new CellReference(workbook.getName("REV").getRefersToFormula());
					assert revRef.getSheetName() != null;
					XSSFRow revRow = workbook.getSheet(revRef.getSheetName()).getRow(revRef.getRow());
					XSSFCell revCell = revRow.getCell(revRef.getCol());
					if(revCell.getNumericCellValue() < 0) {
						//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
						errors += "- REV cell " + revRef.getCellRefParts()[2] + revRef.getCellRefParts()[1] + " in \""
								+ committeeFile.getParentFile().getName() + "\\" + committeeFile.getName() + "\" is not positive\n";
					}
				}

				//Check EXP Name
				if (workbook.getName("EXP") == null) {
					errors += "- EXP cell name missing in \"" + committeeFile.getName() + "\"\n";
				}
				else {
					//Check if Expenses are negative
					CellReference expRef = new CellReference(workbook.getName("EXP").getRefersToFormula());
					assert expRef.getSheetName() != null;
					XSSFRow expRow = workbook.getSheet(expRef.getSheetName()).getRow(expRef.getRow());
					XSSFCell expCell = expRow.getCell(expRef.getCol());
					if(expCell.getNumericCellValue() > 0) {
						//							Column Letter ex "A"		Row number ex "2" (Row is 1 based)
						errors += "- EXP cell " + expRef.getCellRefParts()[2] + expRef.getCellRefParts()[1] + " in \""
								+ committeeFile.getParentFile().getName() + "\\" + committeeFile.getName() + "\" is not negative\n";
					}
				}

				//Check NAME name
				if (workbook.getName("NAME") == null) {
					errors += "- NAME cell name missing in \"" + committeeFile.getName() + "\"\n";
				}

			} catch (Exception e) {
				//Check if file is open
				if (e.getMessage().contains("(The process cannot access the file because it is being used by another process)")) {
					errors += "- Close file: \"" + committeeFile.getParentFile().getName() + "\\" + committeeFile.getName() + "\"\n";
				}
				else {
					//Log all other errors
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
					//Log all other errors
					errors += e.toString();
				}
			}
			if (isCancelled()) {
				updateMessage("Cancelled");
				return false;
			}
		}


		//Check if Previous Budget is open
		if(previousBudgetFile != null) {
			try (XSSFWorkbook workbook = new XSSFWorkbook(previousBudgetFile)) {
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

		//Log any errors if occurred
		if(errors.trim().length() != 0) {
			updateMessage("Errors occurred:\n" + errors);
			return false;
		}
		return true;
	}

	/**
	 * Returns all File objects that qualify as Committee files.
	 * A File qualifies as a Committee file if it is not hidden and has the .xlsx extension
	 *
	 * @param portfolioDirectory the root directory containing Committee Files
	 * @return a File array containing valid committee Files
	 */
	public static File[] getCommitteeFiles(File portfolioDirectory) {
		return portfolioDirectory.listFiles(file -> {
			if (file.isHidden())
				return false;
			else if (file.getName().endsWith(".xlsx"))
				return true;
			return false;
		});
	}

	/**
	 * Returns all File objects that qualify as Portfolio directories.
	 * A File qualifies as a Portfolio if it is not hidden and is a directory
	 *
	 * @param budgetDirectory the root directory containing portfolio subdirectories
	 * @return a File array containing valid portfolio directories
	 */
	static File[] getPortfolioDirectories(File budgetDirectory) {
		return budgetDirectory.listFiles(file -> {
			if(file.isHidden())
				return false;
			else if(file.isDirectory())
				return true;
			return false;
		});
	}

	/**
	 * Returns all qualifying CommitteeFiles in all Portfolios
	 * @param budgetDirectory the root budget directory containing portfolio subdirectories
	 * @return a ArrayList containing all qualifying Committee Files
	 */
	private List<File> getCommitteeFilesToCheck(File budgetDirectory) {

		List<File> committeeFilesToCheck = new ArrayList<>();

		//Get Portfolio directories ignoring files and hidden directories
		File[] portfolioDirectories = getPortfolioDirectories(budgetDirectory);

		assert portfolioDirectories != null;

		//Cycle through portfolio directories getting committee files
		for(File portfolioDirectory : portfolioDirectories) {

			if(isCancelled()) {
				updateMessage("Cancelled");
				return null;
			}

			File[] committeeFiles = getCommitteeFiles(portfolioDirectory);

			if(committeeFiles == null || committeeFiles.length == 0)
				errors += "- Portfolio \"" + portfolioDirectory.getName() + "\" contains no committee budgets\n";
			else
				Collections.addAll(committeeFilesToCheck, committeeFiles);
		}
		return committeeFilesToCheck;
	}
}
