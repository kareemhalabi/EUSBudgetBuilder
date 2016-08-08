package ca.mcgilleus.budgetbuilder.fxml;

import ca.mcgilleus.budgetbuilder.application.BudgetWizzard;
import javafx.event.ActionEvent;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.AnchorPane;
import javafx.stage.Stage;

import java.io.IOException;

/*
 * Created by Kareem Halabi on 2016-08-08.
 */
public class WelcomeController extends AnchorPane{


	public WelcomeController() {
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
		FileSelectController fileSelectController = new FileSelectController();
		BudgetWizzard.primaryStage.setScene(new Scene(fileSelectController));
	}

}
