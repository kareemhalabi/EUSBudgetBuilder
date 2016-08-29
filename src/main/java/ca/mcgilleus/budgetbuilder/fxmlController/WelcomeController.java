package ca.mcgilleus.budgetbuilder.fxmlController;

import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.AnchorPane;

import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizard.primaryStage;

/**
 * This class controls the Welcome Scene where all instructions and disclaimers are presented
 * @author Kareem Halabi
 */
public class WelcomeController extends AnchorPane{

	private static Scene welcomeScene;

	/**
	 * Public getter for Welcome Scene
	 * @return the Welcome Scene instance
	 */
	public static Scene getWelcomeScene() {
		if (welcomeScene == null)
			welcomeScene = new Scene(new WelcomeController());
		return welcomeScene;
	}

	/**
	 * This class and constructor follows a singleton pattern because the welcome scene does not need to
	 * be recreated after switching to another scene and returning
	 */
	private WelcomeController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("/fxml/welcome.fxml"));
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
