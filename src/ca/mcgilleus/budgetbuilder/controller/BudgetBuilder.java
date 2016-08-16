/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.controller;

import ca.mcgilleus.budgetbuilder.fxml.FileSelectController;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import javafx.concurrent.Task;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static ca.mcgilleus.budgetbuilder.controller.PortfolioCreator.*;


public class BudgetBuilder {

	static BuildTask buildTask;
	private static int totalCommitteeFiles;
	static int totalProgress;
	static int currentProgress;
	private static EUSBudget budget;

	public static Task getValidationTask() {

		return new Task() {

			private String errors = "";

			@Override
			protected Boolean call() throws Exception {

				final List<File> filesToCheck = getFilesToCheck(FileSelectController.getSelectedDirectory());

				setTotalCommitteeFiles(filesToCheck.size());

				//filesToCheck may be null due to Task cancellation or some other error
				if (filesToCheck == null)
					return false;

				//Check portfolios for errors
				double currentProgress = 0;
				for (File f : filesToCheck) {

					if(isCancelled()) {
						updateMessage("Cancelled");
						return false;
					}

					try {
						//Check if file is already open
						XSSFWorkbook workbook = new XSSFWorkbook(f);

						//Check names
						if(workbook.getName("AMT") == null) {
							errors += "- AMT cell name missing in \"" + f.getName() + "\"\n";
						}
						if(workbook.getName("COMM_NAME") == null) {
							errors += "- COMM_NAME cell name missing in \"" + f.getName() + "\"\n";
						}

						workbook.close();
					} catch (Exception e) {
						if(e.getMessage().contains("(The process cannot access the file because it is being used by another process)"))
							errors += "- Close file: \"" + f.getName() + "\"\n";
						else
							errors += e.toString();
					}
					updateProgress(++currentProgress, filesToCheck.size());
				}

				if(isCancelled()) {
					updateMessage("Cancelled");
					return false;
				}

				//Check if outputFile is open
				File outputFile = FileSelectController.getOutputFile();
				if(outputFile.exists()) {
					try {
						//Don't understand why this is necessary, prevents a Zip bomb error
						ZipSecureFile.setMinInflateRatio(0.001);
						XSSFWorkbook workbook = new XSSFWorkbook(outputFile);
						workbook.close();
					} catch (Exception e) {
						if(e.getMessage().contains("(The process cannot access the file because it is being used by another process)"))
							errors += "- Close file: \"" + outputFile.getName() + "\"\n";
						else
							errors += e.toString();
					}
				}

				if(errors.trim().length() != 0) {
					updateMessage("Errors occurred:\n" + errors);
					return false;
				}
				return true;
			}

			private List<File> getFilesToCheck(File rootDirectory) {

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

			XSSFSheet overviewSheet = budget.getWb().createSheet(budget.getBudgetYear() + " Budget");

			File[] portfolioDirectories = getPortfolioDirectories(FileSelectController.getSelectedDirectory());

			//Total progress: total committee sheets + each portfolio overview + budget overview
			totalProgress = totalCommitteeFiles + portfolioDirectories.length + 1;
			currentProgress = 0;

			//Cycle through each portfolio folder
			for(File portfolioDirectory : portfolioDirectories) {
				if(isCancelled()) {
					updateBuildMessage("Cancelled");
					return false;
				}
				PortfolioCreator.createPortfolio(portfolioDirectory, budget);
			}

			updateBuildMessage("Compiling EUS Budget Overview");
			updateProgress(++currentProgress, totalProgress);

			writeHeader(budget, overviewSheet);
			for(Portfolio p : budget.getPortfolios())
				createPortfolioOverview(p, overviewSheet);

			writeTotals(overviewSheet);

			overviewSheet.createFreezePane(0,1);

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
	 * Creates a new budget based on Fall Semester year
	 * @return a new budget based on Fall Semester year
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

	public static void setTotalCommitteeFiles(int totalCommitteeFiles1) {
		totalCommitteeFiles = totalCommitteeFiles1;
	}
}