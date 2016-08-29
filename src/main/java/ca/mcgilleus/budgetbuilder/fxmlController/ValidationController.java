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

/**
 * This class controls the Validation Scene
 *
 * @author Kareem Halabi
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

	/**
	 * Creates controller instance for validate.fxml
	 * When a new ValidationController is created, a new ValidationTask is also created. The validation task's progress
	 * property is bound to the progress bar and a changeListener for the message property updates
	 * the validationConsole textArea. ValidationThread is set statically so that it can be accessed in BudgetWizard
	 * in the case the application closes during this step.
	 * <p>
	 * Unlike WelcomeController and FileSelectController, ValidationController is not singleton because if the user
	 * cancels the task or if it fails, a new ValidationController and ValidationTask should be re-created when the user
	 * retries this step
	 *
	 * @see ca.mcgilleus.budgetbuilder.application.BudgetWizard
	 * @see WelcomeController
	 * @see FileSelectController
	 */
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

		validationTask.messageProperty().addListener((observable, oldValue, newValue) ->
			validationConsole.appendText(newValue + "\n")
		);

		//Set on fail not working so returned boolean value determines if task succeeded
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

	/**
	 * Defines the behaviour when the ValidationTask fails or is cancelled.
	 * Progress bar is reset to 0 and Cancel button switches to a back button to return to File Select
	 * Retry-button re-generates a new ValidationController to restart this step
	 */
	private void cancelOrFail() {
		validationProgressBar.progressProperty().unbind();
		validationProgressBar.setProgress(0);

		cancelBackBtn.setText("Back");
		cancelBackBtn.setOnAction(event1 -> showFileSelect());

		retryBtn.setDisable(false);
		retryBtn.setOnAction(event1 ->
			primaryStage.setScene(new Scene(new ValidationController()))
		);
	}
}
