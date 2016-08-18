/**
 * © Kareem Halabi 2016
 * @author Kareem Halabi
 */

package ca.mcgilleus.budgetbuilder.controller;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.controller.BudgetBuilder.*;

public class CommitteeCreator {

	public static void createCommitteeBudget(IndexedColors color, Portfolio p, File committeeFile) {

		try (XSSFWorkbook committeeWB = new XSSFWorkbook(committeeFile)) {
			XSSFSheet committeeSheet = committeeWB.getSheetAt(0);

			Name name = committeeWB.getName("NAME");
			CellReference nameRef = new CellReference(name.getRefersToFormula());
			XSSFRow nameRow = committeeSheet.getRow(nameRef.getRow());
			XSSFCell nameCell = nameRow.getCell(nameRef.getCol());
			String sheetName = nameCell.getStringCellValue();

			//TODO throw error instead
			// Get full name if abbreviated unavailable
			if (sheetName == null || sheetName.trim().length() == 0) {
				nameRow = committeeSheet.getRow(nameRef.getRow() - 1);
				nameCell = nameRow.getCell(nameRef.getCol());
				sheetName = nameCell.getStringCellValue();
			}

			// So that AMT reference has correct sheet name
			committeeWB.setSheetName(0, sheetName);

			Name revRef = committeeWB.getName("REV");
			Name expRef = committeeWB.getName("EXP");

			new CommitteeBudget(sheetName, revRef.getRefersToFormula(), expRef.getRefersToFormula(), p);

			XSSFSheet bSheet = BudgetBuilder.getWorkbook().createSheet(sheetName);
			Cloner.cloneSheet(committeeSheet, bSheet);

			bSheet.setTabColor(color.getIndex());

			buildTask.updateBuildMessage("\t -" + sheetName);
			buildTask.updateBuildProgress(++currentProgress, totalProgress);
		} catch (IOException | InvalidFormatException e) {
			buildTask.updateBuildMessage(e.toString());
		}
	}

}
