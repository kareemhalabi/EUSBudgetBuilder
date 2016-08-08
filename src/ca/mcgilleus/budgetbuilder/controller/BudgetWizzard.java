package ca.mcgilleus.budgetbuilder.controller;
/**
 * Created by Kareem Halabi on 2016-08-07.
 */

import javafx.application.Application;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import java.io.File;
import java.io.IOException;

public class BudgetWizzard extends Application {

	private static Stage stage;

	public TextField folderTextField;
	public TextField outputFileTextField;
	@FXML
	public Button fileSelectNextBtn;
	private File outputFile;
	private File selectedDirectory;

	public static void main(String[] args) { launch(args);	}

	public Scene getFXMLScene(String URI, int width, int height) {
		Parent root = null;
		try {
			root = FXMLLoader.load(getClass().getResource(URI));
		} catch (IOException e) {
			e.printStackTrace();
		}
		return new Scene(root, width, height);
	}

	@Override
	public void start(Stage primaryStage) throws IOException {
		stage = primaryStage;
		stage.setTitle("EUS Budget Builder");
		stage.setResizable(false);
		stage.getIcons().add(new Image("ca/mcgilleus/budgetbuilder/EUSfavicon.png"));
		showWelcomeScene();
	}

	public void showWelcomeScene() {
		stage.setScene(getFXMLScene("../fxml/welcome.fxml",500,300));
		stage.show();
	}

	//------------------- FileSelectScene -------------------
	public void showFileSelectScene() {
		stage.setScene(getFXMLScene("../fxml/file_select.fxml",500,300));
//		if(selectedDirectory == null || outputFile == null)
//			fileSelectNextBtn.setDisable(true);
		stage.show();
	}

	public void showDirectoryChooser() {
		DirectoryChooser directoryChooser = new DirectoryChooser();
		directoryChooser.setTitle("Select root folder");
		selectedDirectory = directoryChooser.showDialog(stage);
		if (selectedDirectory != null)
			folderTextField.setText(selectedDirectory.getPath());
	}

	public void showFileSaveChooser() {
		FileChooser fileSaveChooser = new FileChooser();
		fileSaveChooser.setTitle("Save output as");
		fileSaveChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
		outputFile = fileSaveChooser.showSaveDialog(stage);
		if (outputFile != null)
			outputFileTextField.setText(outputFile.getPath());
	}
	//-------------------------------------------------------

	public void showValidateScene() {

	}
}
