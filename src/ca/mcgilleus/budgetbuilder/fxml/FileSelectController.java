package ca.mcgilleus.budgetbuilder.fxml;

import ca.mcgilleus.budgetbuilder.application.BudgetWizzard;
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

import static ca.mcgilleus.budgetbuilder.application.BudgetWizzard.primaryStage;

/*
 * Created by Kareem Halabi on 2016-08-08.
 */
public class FileSelectController extends AnchorPane{

	private static Scene fileSelectScene;
	private static FileSelectController theInstance;

	@FXML
	public TextField folderTextField;
	@FXML
	public TextField outputFileTextField;
	@FXML
	public Button fileSelectNextBtn;

	private File outputFile;
	private File selectedDirectory;

	public static Scene getFileSelectScene() {
		if (fileSelectScene == null)
			fileSelectScene = new Scene(getInstance());

		getInstance().checkNext();
		return fileSelectScene;
	}

	public static FileSelectController getInstance() {
		if(theInstance == null)
			theInstance = new FileSelectController();
		return theInstance;
	}

	private FileSelectController() {
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("file_select.fxml"));
		fxmlLoader.setRoot(this);
		fxmlLoader.setController(this);

		try {
			fxmlLoader.load();
		} catch (IOException exception) {
			throw new RuntimeException(exception);
		}
	}

	public void checkNext() {
		if(selectedDirectory == null || outputFile == null)
			fileSelectNextBtn.setDisable(true);
		else {
			fileSelectNextBtn.setDisable(false);
		}
	}

//	public void initialize() {
//		if(selectedDirectory == null || outputFile == null)
//			fileSelectNextBtn.setDisable(true);
//		else {
//			fileSelectNextBtn.setDisable(false);
//		}
//	}

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

	public void showValidateScene() {

	}

	public void showWelcomeScene() {
		primaryStage.setScene(WelcomeController.getWelcomeScene());
	}
}
