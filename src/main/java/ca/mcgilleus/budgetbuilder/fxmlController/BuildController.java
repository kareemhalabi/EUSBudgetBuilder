package ca.mcgilleus.budgetbuilder.fxmlController;

import ca.mcgilleus.budgetbuilder.task.BuildTask;
import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;
import javafx.scene.layout.AnchorPane;

import java.awt.*;
import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizard.primaryStage;
import static ca.mcgilleus.budgetbuilder.fxmlController.FileSelectController.outputFile;

/**
 * This class controls the Build Scene
 *
 * @author Kareem Halabi
 */
public class BuildController extends AnchorPane{

	@FXML
	public Button finishBtn;
	@FXML
	public Button cancelBackBtn;
	@FXML
	public ProgressBar buildProgressBar;
	@FXML
	public TextArea buildConsole;

	private Task buildTask;
	public static Thread buildThread;

	/**
	 * Creates controller instance for build.fxml
	 * When a new BuildController is created, a new BuildTask is also created. The build task's progress
	 * property is bound to the progress bar and a changeListener for the message property updates
	 * the buildConsole textArea. BuildThread is set statically so that it can be accessed in BudgetWizard
	 * in the case the application closes during this step.
	 * <p>
	 * Unlike WelcomeController and FileSelectController, BuildController is not singleton because if the user
	 * cancels the task or if it fails, a new BuildController and BuildTask should be re-created when the user retries
	 * this step
	 *
	 * @see ca.mcgilleus.budgetbuilder.application.BudgetWizard
	 * @see WelcomeController
	 * @see FileSelectController
	 */
	BuildController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("/fxml/build.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}

		buildTask = new BuildTask();

		buildProgressBar.progressProperty().unbind();
		buildProgressBar.progressProperty().bind(buildTask.progressProperty());

		cancelBackBtn.setOnAction(event -> {
			buildTask.cancel(false);
			cancelOrFail();
		});

		buildTask.messageProperty().addListener((observable, oldValue, newValue) -> {
			buildConsole.appendText(newValue +"\n");
		});

		//Set on fail not working so returned boolean value determines if task succeeded
		buildTask.setOnSucceeded(event -> {
			if((boolean)buildTask.getValue()) {
				cancelBackBtn.setDisable(true);
				finishBtn.setDisable(false);
				finishBtn.setOnAction(event1 -> {
					try {
						primaryStage.hide();
						Desktop.getDesktop().open(outputFile);
						Platform.exit();
						System.exit(0);
					} catch (IOException e) {
						buildConsole.appendText(e.toString());
					}
				});
			}
			else cancelOrFail();
		});

		buildThread = new Thread(buildTask, "Build Thread");
		buildThread.start();
	}

	/**
	 * Defines the behaviour when the Build task fails or is cancelled.
	 */
	private void cancelOrFail() {
		buildProgressBar.progressProperty().unbind();
		buildProgressBar.setProgress(0);

		cancelBackBtn.setText("Back");
		cancelBackBtn.setOnAction(event -> showFileSelect());
	}

	/**
	 * Returns to FileSelect Scene
	 *
	 * @see FileSelectController
	 */
	private void showFileSelect() {
		primaryStage.setScene(FileSelectController.getFileSelectScene());
	}
}
