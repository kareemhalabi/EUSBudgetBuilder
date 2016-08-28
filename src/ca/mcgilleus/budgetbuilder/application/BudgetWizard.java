package ca.mcgilleus.budgetbuilder.application;
/*
 * Created by Kareem Halabi on 2016-08-07.
 */

import ca.mcgilleus.budgetbuilder.fxmlController.WelcomeController;
import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.fxmlController.BuildController.buildThread;
import static ca.mcgilleus.budgetbuilder.fxmlController.ValidationController.validationTask;
import static ca.mcgilleus.budgetbuilder.fxmlController.ValidationController.validationThread;
import static ca.mcgilleus.budgetbuilder.service.BudgetBuilder.buildTask;

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

	@Override
	public void stop() {
		cancelTask(validationTask, validationThread);
		cancelTask(buildTask, buildThread);
	}

	private void cancelTask(Task task, Thread thread) {
		if(task != null && thread != null) {
			if(task.isRunning()) {
				task.cancel(false);
				try {
					thread.join();
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
