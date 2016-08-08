package ca.mcgilleus.budgetbuilder.fxml;

import ca.mcgilleus.budgetbuilder.application.BudgetWizzard;
import javafx.event.ActionEvent;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.AnchorPane;
import javafx.stage.Stage;

import java.io.File;
import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizzard.primaryStage;

/*
 * Created by Kareem Halabi on 2016-08-08.
 */
public class WelcomeController extends AnchorPane{

	private static WelcomeController theInstance;
	private static Scene welcomeScene;

	public static Scene getWelcomeScene() {
		if (welcomeScene == null)
			welcomeScene = new Scene(getInstance());
		return welcomeScene;
	}

	public static WelcomeController getInstance() {
		if (theInstance == null)
			theInstance = new WelcomeController();

		return theInstance;
	}

	private WelcomeController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("welcome.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}
	}

	public void showFileSelectScene(ActionEvent actionEvent) {
		primaryStage.setScene(FileSelectController.getFileSelectScene());
	}

}
