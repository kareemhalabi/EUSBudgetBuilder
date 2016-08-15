/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.controller;

import ca.mcgilleus.budgetbuilder.fxml.FileSelectController;
import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Styles;
import javafx.concurrent.Task;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.List;
import java.util.Queue;

public class BudgetBuilder {

	static BuildTask buildTask;
	private static int totalCommitteeFiles;
	public static void setTotalCommitteeFiles(int totalCommitteeFiles1) {
		totalCommitteeFiles = totalCommitteeFiles1;
	}

	static int totalProgress;
	static int currentProgress;

	private static XSSFWorkbook wb;

	public static XSSFWorkbook getWorkbook() {
		return wb;
	}

	public static void main(String[] args) {

		File root = new File(System.getProperty("user.dir") + File.separator + "W2017 Budget");
		EUSBudget budget = createBudget();
		wb = budget.getWorkbook();

		wb.createSheet(budget.getBudgetYear() + " Budget");

		Queue<IndexedColors> availableTabColors = new LinkedList<>();
		createColorQueue(availableTabColors);

		PortfolioCreator.setAllColors(availableTabColors);

		for(File f : root.listFiles()) {
			PortfolioCreator.createPortfolio(f, budget);
		}
System.out.println("Compiling EUS Budget Overview");
		createEUSBudgetOverview(budget);

		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(root.getName() + ".xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();

			System.out.println("Openning: " + root.getName() + ".xlsx");
			Desktop.getDesktop().open(new File(root.getName() + ".xlsx"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static void createEUSBudgetOverview(EUSBudget budget) {

		XSSFSheet overviewSheet = wb.getSheet(budget.getBudgetYear() + " Budget");
		XSSFRow header = overviewSheet.createRow(0);

		XSSFCell title = header.createCell(0, Cell.CELL_TYPE_STRING);
		title.setCellValue("Portfolio");

		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		title.setCellValue("Committee");

		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);

		title.setCellValue(budget.getBudgetYear() + " Budget");

		for(int i = 0; i < budget.getPortfolios().size(); i++) {
			Portfolio p = budget.getPortfolio(i);
			for(int j = 0; j < p.getCommitteeBudgets().size(); j++) {
				CommitteeBudget committee = p.getCommitteeBudget(j);
				XSSFRow row = overviewSheet.createRow(overviewSheet.getLastRowNum()+1);

				XSSFCell portfolio = row.createCell(0, Cell.CELL_TYPE_STRING);
				portfolio.setCellValue(p.getName());
				portfolio.setCellStyle(p.getPortfolioLabelStyle());

				XSSFCell committeeName = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
				committeeName.setCellValue(committee.getName());

				XSSFCell committeeAmt = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_FORMULA);
				committeeAmt.setCellFormula(committee.getAmtRequestedRef());
			}
		}

		writeTotals(overviewSheet);

		styleBudgetOverview(overviewSheet);

		overviewSheet.createFreezePane(0,1);
	}

	private static void styleBudgetOverview(XSSFSheet overviewSheet) {

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

		for(int j = 1; j < overviewSheet.getLastRowNum()-1; j++) {
			XSSFRow dataRow = overviewSheet.getRow(j);

			//Portfolio label styling is done separately in createEUSBudgetOverview

			dataRow.getCell(1).setCellStyle(Styles.COMMITTEE_LABEL_STYLE);

			for (int k = 2; k < dataRow.getLastCellNum(); k++)
				dataRow.getCell(k).setCellStyle(Styles.CURRENCY_CELL_STYLE);
		}

	}

	public static void createColorQueue(Queue<IndexedColors> colors) {
		colors.add(IndexedColors.MAROON);
		colors.add(IndexedColors.LIGHT_ORANGE);
		colors.add(IndexedColors.YELLOW);
		colors.add(IndexedColors.GREEN);
		colors.add(IndexedColors.BLUE);
		colors.add(IndexedColors.PLUM);
		colors.add(IndexedColors.GREY_40_PERCENT);

	}

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

		//Styling

		totalLabel.setCellStyle(Styles.TOTAL_LABEL_STYLE);

		for(int l = 2; l < totals.getLastCellNum(); l++)
			totals.getCell(l).setCellStyle(Styles.TOTAL_CELL_STYLE);
	}

	public static EUSBudget createBudget() {
		Calendar c = Calendar.getInstance();
		Date budgetYear;
		if(c.get(Calendar.MONTH) <= Calendar.FEBRUARY)
			c.add(Calendar.YEAR, -1);

		budgetYear = c.getTime();

		return new EUSBudget(budgetYear);
	}

	private static File[] getPortfolioDirectories(File rootDirectory) {
		return rootDirectory.listFiles(file -> {
			if(file.isHidden())
				return false;
			else if(file.isDirectory())
				return true;
			return false;
		});
	}

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

			EUSBudget budget = createBudget();
			wb = budget.getWorkbook();

			wb.createSheet(budget.getBudgetYear() + " Budget");

			Queue<IndexedColors> availableTabColors = new LinkedList<>();
			createColorQueue(availableTabColors);

			PortfolioCreator.setAllColors(availableTabColors);

			File[] portfolioDirectories = getPortfolioDirectories(FileSelectController.getSelectedDirectory());

			//Total progress: total committee sheets + each portfolio overview + budget overview
			totalProgress = totalCommitteeFiles + portfolioDirectories.length + 1;
			currentProgress = 0;

			for(File portfolioDirectory : portfolioDirectories) {
				if(isCancelled()) {
					updateBuildMessage("Cancelled");
					return false;
				}
				PortfolioCreator.createPortfolio(portfolioDirectory, budget);
			}

			updateBuildMessage("Compiling EUS Budget Overview");
			updateProgress(++currentProgress, totalProgress);
			createEUSBudgetOverview(budget);

			FileOutputStream fileOut;
			try {
				fileOut = new FileOutputStream(FileSelectController.getOutputFile());
				wb.write(fileOut);
				fileOut.close();
				wb.close();

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
}
