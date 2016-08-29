package ca.mcgilleus.budgetbuilder.fxmlController;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.AnchorPane;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.File;
import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.application.BudgetWizard.primaryStage;

/**
 * This class controls the File Select Scene where all inputs are set for validation and budget compilation
 * @author Kareem Halabi
 */
public class FileSelectController extends AnchorPane{

	private static Scene fileSelectScene;

	@FXML
	public TextField folderTextField;
	@FXML
	public TextField outputFileTextField;
	@FXML
	public TextField previousFileTextField;
	@FXML
	public TextField previousNameTextField;
	@FXML
	public Button fileSelectNextBtn;

	public static File selectedDirectory;
	public static File outputFile;
	public static File previousBudgetFile;

	public static String previousBudgetName;

	/**
	 * Public getter for File Select Scene
	 * @return the File Select Scene instance
	 */
	static Scene getFileSelectScene() {
		if (fileSelectScene == null)
			fileSelectScene = new Scene(new FileSelectController());
		return fileSelectScene;
	}

	/**
	 * This class and constructor follows a singleton pattern because the text in all TextFields needs to
	 * persist after switching to another scene and returning
	 */
	private FileSelectController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("/fxml/file_select.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}

		previousNameTextField.textProperty().addListener((observable, oldValue, newValue) -> checkNext());

		checkNext();
	}

	public void showDirectoryChooser() {
		DirectoryChooser directoryChooser = new DirectoryChooser();
		directoryChooser.setTitle("Select root folder");
		selectedDirectory = directoryChooser.showDialog(primaryStage);
		if (selectedDirectory != null) {
			folderTextField.setText(selectedDirectory.getPath());
			checkNext();
		}
	}

	public void showFileSaveChooser() {
		FileChooser fileSaveChooser = new FileChooser();
		fileSaveChooser.setTitle("Save output as");
		fileSaveChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
		outputFile = fileSaveChooser.showSaveDialog(primaryStage);
		if (outputFile != null) {
			outputFileTextField.setText(outputFile.getPath());
			checkNext();
		}
	}

	public void showFileSelectChooser() {
		FileChooser fileSelectChooser = new FileChooser();
		fileSelectChooser.setTitle("Select previous budget");
		fileSelectChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
		previousBudgetFile = fileSelectChooser.showOpenDialog(primaryStage);
		if(previousBudgetFile != null) {
			previousFileTextField.setText(previousBudgetFile.getPath());
			previousNameTextField.setEditable(true);
		}
		checkNext();
	}

	/**
	 * Clears the previous budget fields
	 */
	public void clearFile() {
		previousBudgetFile = null;
		previousBudgetName = null;
		previousFileTextField.clear();
		previousNameTextField.clear();
		previousNameTextField.setEditable(false);
		checkNext();
	}

	/**
	 * Checks if the next button can be enabled. The next button is enabled when there is a selected directory
	 * and an output file. If a previous budget is selected, it must have a previous name.
	 */
	private void checkNext() {
		if(selectedDirectory == null || outputFile == null ||
				(previousBudgetFile != null && previousNameTextField.getText().isEmpty())) {
			fileSelectNextBtn.setDisable(true);
		} else {
			fileSelectNextBtn.setDisable(false);
		}
	}

	public void showValidateScene() {
		previousBudgetName = previousNameTextField.getText();
		primaryStage.setScene(new Scene(new ValidationController()));
	}

	public void showWelcomeScene() {
		primaryStage.setScene(WelcomeController.getWelcomeScene());
	}
}
