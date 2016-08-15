package ca.mcgilleus.budgetbuilder.fxml;

import ca.mcgilleus.budgetbuilder.controller.BudgetBuilder;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;
import javafx.scene.layout.AnchorPane;

import java.awt.*;
import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizzard.primaryStage;

/*
 * Created by Kareem Halabi on 2016-08-08.
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

	private BudgetBuilder.BuildTask buildTask;

	// Don't want BuildController object to be re-used, so a new one will be
	// created each time build scene is set
	BuildController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("build.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}

		buildTask = new BudgetBuilder.BuildTask();

		buildProgressBar.progressProperty().unbind();
		buildProgressBar.progressProperty().bind(buildTask.progressProperty());

		cancelBackBtn.setOnAction(event -> {
			buildTask.cancel(true);
			cancelOrFail();
		});

		buildTask.messageProperty().addListener((observable, oldValue, newValue) -> {
			buildConsole.appendText(newValue +"\n");
		});

		buildTask.setOnSucceeded(event -> {
			if((boolean)buildTask.getValue()) {
				finishBtn.setDisable(false);
				finishBtn.setOnAction(event1 -> {
					try {
						primaryStage.hide();
						Desktop.getDesktop().open(FileSelectController.getOutputFile());
						Platform.exit();
						System.exit(0);
					} catch (IOException e) {
						buildConsole.appendText(e.toString());
					}
				});
			}
			else cancelOrFail();
		});

		new Thread(buildTask).start();
	}

	private void cancelOrFail() {
		buildProgressBar.progressProperty().unbind();
		buildProgressBar.setProgress(0);

		cancelBackBtn.setText("Back");
		cancelBackBtn.setOnAction(event -> showFileSelect());
	}

	private void showFileSelect() {
		primaryStage.setScene(FileSelectController.getFileSelectScene());
	}

}
