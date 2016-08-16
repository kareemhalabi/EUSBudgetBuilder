package ca.mcgilleus.budgetbuilder.application;
/*
 * Created by Kareem Halabi on 2016-08-07.
 */

import ca.mcgilleus.budgetbuilder.fxml.WelcomeController;
import javafx.application.Application;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.IOException;

public class BudgetWizard extends Application {

	public static Stage primaryStage;

	public static void main(String[] args) { launch(args);	}

	@Override
	public void start(Stage stage) throws IOException {
		primaryStage = stage;
		primaryStage.setTitle("EUS Budget Builder");
		primaryStage.setResizable(false);
		primaryStage.getIcons().add(new Image("ca/mcgilleus/budgetbuilder/EUSfavicon.png"));
		primaryStage.setScene(WelcomeController.getWelcomeScene());
		primaryStage.sizeToScene();
		primaryStage.show();
	}
}
