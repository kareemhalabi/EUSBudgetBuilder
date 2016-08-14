package ca.mcgilleus.budgetbuilder.fxml;

import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;
import javafx.scene.layout.AnchorPane;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizzard.primaryStage;

/*
 * Created by Kareem Halabi on 2016-08-08.
 */
public class ValidationController extends AnchorPane{

	@FXML
	public Button retryBtn;
	@FXML
	public Button cancelBackBtn;
	@FXML
	public ProgressBar validationProgressBar;
	public TextArea validationConsole;


	private ValidationTask validationTask;

	private File outputFile;
	private File selectedDirectory;

	// Don't want ValidationController object to be re-used, so a new one will be
	// created each time validation scene is set
	public ValidationController(File selectedDirectory, File outputFile) {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("validate.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}

		this.selectedDirectory = selectedDirectory;
		this.outputFile = outputFile;

		validationTask = new ValidationTask(selectedDirectory, outputFile);

		validationProgressBar.progressProperty().unbind();
		validationProgressBar.progressProperty().bind(validationTask.progressProperty());

		cancelBackBtn.setOnAction(event -> {
			validationTask.cancel(true);
			validationConsole.setText("Cancelled\n");
			cancelOrFail();
		});

		validationTask.messageProperty().addListener((observable, oldValue, newValue) -> {
			validationConsole.appendText(newValue + "\n");
		});

		//Task setOnFail is not working, so fail status is checked here
		validationTask.setOnSucceeded(event -> {
			if ((boolean)validationTask.getValue()) {
				//TODO, proceed to Building
			}
			else {
				cancelOrFail();
			}
		});

		new Thread(validationTask).start();
	}

	private void showFileSelect() {
		primaryStage.setScene(FileSelectController.getFileSelectScene());
	}

	private void cancelOrFail() {
		validationProgressBar.progressProperty().unbind();
		validationProgressBar.setProgress(0);

		cancelBackBtn.setText("Back");
		cancelBackBtn.setOnAction(event1 -> showFileSelect());

		retryBtn.setDisable(false);
		retryBtn.setOnAction(event1 -> {
			primaryStage.setScene(new Scene(new ValidationController(selectedDirectory, outputFile)));
		});
	}

	public static class ValidationTask extends Task {

		private File selectedDirectory;
		private File outputFile;

		private String errors = "";

		public ValidationTask(File selectedDirectory, File outputFile) {
			this.selectedDirectory = selectedDirectory;
			this.outputFile = outputFile;
		}

		@Override
		protected Boolean call() throws Exception {

			final List<File> filesToCheck = getFilesToCheck(selectedDirectory);

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
					XSSFWorkbook workbook = new XSSFWorkbook(f);
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
						errors += e.getMessage();
				}
				updateProgress(++currentProgress, filesToCheck.size());
			}

			if(isCancelled()) {
				updateMessage("Cancelled");
				return false;
			}

			//Check if outputFile is open
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
//			if(!rootDirectory.isDirectory()) {
//				errors += "- Selected file is not a directory\n";
//				return null;
//			}

			ArrayList<File> filesToCheck = new ArrayList<>();

			//Get Portfolio directories ignoring files and hidden directories
			File[] subDirectories = rootDirectory.listFiles(file -> {
				if(file.isHidden())
					return false;
				else if(file.isDirectory())
					return true;
				return false;
			});

			assert subDirectories != null;
			//Get Committee excel files ignoring any hidden files
			for(File subDirectory : subDirectories) {

				if(isCancelled()) {
					updateMessage("Cancelled");
					return null;
				}

				//noinspection ConstantConditions
				File[] files = subDirectory.listFiles(file -> {
					if (file.isHidden())
						return false;
					else if (file.getName().endsWith(".xlsx"))
						return true;
					return false;
				});

				if(files.length == 0)
					errors += "- Portfolio \"" + subDirectory.getName() + "\" contains no committee budgets\n";
				Collections.addAll(filesToCheck, files);
			}
			return filesToCheck;
		}

	}
}
