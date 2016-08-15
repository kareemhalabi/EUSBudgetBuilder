package ca.mcgilleus.budgetbuilder.fxml;

import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.AnchorPane;

import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizzard.primaryStage;

/*
 * Created by Kareem Halabi on 2016-08-08.
 */
public class WelcomeController extends AnchorPane{

	private static Scene welcomeScene;

	public static Scene getWelcomeScene() {
		if (welcomeScene == null)
			welcomeScene = new Scene(new WelcomeController());
		return welcomeScene;
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

	public void showFileSelectScene() {
		primaryStage.setScene(FileSelectController.getFileSelectScene());
	}

}
