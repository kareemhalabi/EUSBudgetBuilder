package ca.mcgilleus.budgetbuilder.fxmlController;

import ca.mcgilleus.budgetbuilder.task.ValidationTask;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;
import javafx.scene.layout.AnchorPane;

import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizard.primaryStage;

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
	@FXML
	public TextArea validationConsole;

	public static Task validationTask;
	public static Thread validationThread;

	// Don't want ValidationController object to be re-used, so a new one will be
	// created each time validation scene is set
	ValidationController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("/fxml/validate.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}

		validationTask = new ValidationTask();

		validationProgressBar.progressProperty().unbind();
		validationProgressBar.progressProperty().bind(validationTask.progressProperty());

		cancelBackBtn.setOnAction(event -> {
			validationTask.cancel(false);
			cancelOrFail();
		});

		validationTask.messageProperty().addListener((observable, oldValue, newValue) -> {
			validationConsole.appendText(newValue + "\n");
		});

		//Task setOnFail is not working, so fail status is checked here
		validationTask.setOnSucceeded(event -> {
			if ((boolean)validationTask.getValue()) {
				primaryStage.setScene(new Scene(new BuildController()));
			}
			else cancelOrFail();
		});

		validationThread = new Thread(validationTask, "Validation Thread");
		validationThread.start();
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
			primaryStage.setScene(new Scene(new ValidationController()));
		});
	}
}
