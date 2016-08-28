package ca.mcgilleus.budgetbuilder.task;

import ca.mcgilleus.budgetbuilder.service.BudgetBuilder;
import ca.mcgilleus.budgetbuilder.service.PortfolioCreator;
import javafx.concurrent.Task;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.*;
import static ca.mcgilleus.budgetbuilder.task.ValidationTask.getCommitteeFiles;
import static ca.mcgilleus.budgetbuilder.util.Styles.initStyles;

/**
 * This Task performs the necessary operations for building the budget. First the current task is set, a new EUSBudget instance is created,
 * and styles are initialized. If a previous budget file exists, it is read and rebuilt into model instances. Then
 * each Misc Portfolio Overview is created (alphabetically) followed by each Portfolio Overview. Finally the master
 * overview is created and the new file is written out
 */
public class BuildTask extends Task {

	@Override
	protected Boolean call() throws Exception {

		BudgetBuilder.buildTask = this;

		BudgetBuilder.budget = BudgetBuilder.createBudget();

		initStyles();

		if (previousBudgetFile != null)
			BudgetBuilder.rebuildPreviousBudget();
		if (isCancelled()) {
			updateBuildMessage("Cancelled");
			return false;
		}

		XSSFSheet budgetOverview = BudgetBuilder.budget.getWb().createSheet(BudgetBuilder.budget.getBudgetYear() + " Budget");

		File[] miscPortfolioFiles = getCommitteeFiles(selectedDirectory);
		Arrays.sort(miscPortfolioFiles);

		File[] portfolioDirectories = ValidationTask.getPortfolioDirectories(selectedDirectory);
		Arrays.sort(portfolioDirectories);

		//Total progress: total misc portfolios + total committee sheets + each portfolio overview + budget overview
		BudgetBuilder.totalProgress = BudgetBuilder.totalMiscPortfolios + BudgetBuilder.totalCommitteeFiles + portfolioDirectories.length + 1;
		BudgetBuilder.currentProgress = 0;

		//Cycle through each misc portfolio
		for (File miscPortfolio : miscPortfolioFiles) {
			PortfolioCreator.createMiscPortfolio(miscPortfolio);
			if (isCancelled()) {
				updateBuildMessage("Cancelled");
				return false;
			}
		}

		//Cycle through each portfolio folder
		for (File portfolioDirectory : portfolioDirectories) {
			PortfolioCreator.createPortfolio(portfolioDirectory);
			if (isCancelled()) {
				updateBuildMessage("Cancelled");
				return false;
			}
		}

		BudgetBuilder.createBudgetOverview(budgetOverview);

		//Write out finished file
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(outputFile);
			BudgetBuilder.budget.getWb().write(fileOut);
			fileOut.close();
			BudgetBuilder.budget.getWb().close();
			updateBuildMessage("Done! Press finish to open");
		} catch (IOException e) {
			updateBuildMessage(e.toString());
		}

		return true;
	}

	/**
	 * Provides public access to javafx.concurrent.Task updateMessage()
	 * @see javafx.concurrent.Task
	 * @param message the message to update
	 */
	public void updateBuildMessage(String message) {
		updateMessage(message);
	}

	/**
	 * Provides public access to javafx.concurrent.Task updateProgress()
	 * @see javafx.concurrent.Task
	 * @param workDone the current work done
	 * @param max the max work
	 */
	public void updateBuildProgress(double workDone, double max) {
		updateProgress(workDone, max);
	}
}
