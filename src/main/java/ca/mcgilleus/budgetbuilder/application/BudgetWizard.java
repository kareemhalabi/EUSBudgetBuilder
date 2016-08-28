package ca.mcgilleus.budgetbuilder.application;

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

/**
 * Entry point for BudgetWizard application. Initializes stage and contains on close listener
 *
 * @author Kareem Halabi
 */
public class BudgetWizard extends Application {

	public static Stage primaryStage;

	public static void main(String[] args) { launch(args);	}

	@Override
	public void start(Stage stage) throws IOException {
		primaryStage = stage;
		primaryStage.setTitle("EUS Budget Builder");
		primaryStage.setResizable(false);
		primaryStage.getIcons().add(new Image("EUSfavicon.png"));
		primaryStage.setScene(WelcomeController.getWelcomeScene());
		primaryStage.sizeToScene();
		primaryStage.show();
	}

	@Override
	public void stop() {
		cancelTask(validationTask, validationThread);
		cancelTask(buildTask, buildThread);
	}

	/**
	 * Safely cancels a task, if running.
	 * @param task the task to cancel. If null, this method does nothing
	 * @param thread the thread containing the task. Will be joined to this thread to allow
	 *               the task to safely close. If null, this method does nothing.
	 * @see ca.mcgilleus.budgetbuilder.task.ValidationTask
	 * @see ca.mcgilleus.budgetbuilder.fxmlController.ValidationController
	 * @see ca.mcgilleus.budgetbuilder.task.BuildTask
	 * @see ca.mcgilleus.budgetbuilder.fxmlController.BuildController
	 */
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
