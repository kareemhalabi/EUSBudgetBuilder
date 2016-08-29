package ca.mcgilleus.budgetbuilder.service;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Cloner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static ca.mcgilleus.budgetbuilder.service.BudgetBuilder.*;

/**
 * Sub-service for creating committees and cloning sheets over to the output file
 *
 * @author Kareem Halabi
 */
class CommitteeCreator {

	/**
	 * Creates a new Committee and Committee Sheet
	 * <p>
	 *  First the Committee Name, Revenues and Expenses are obtained from cell names in the committee workbook file.
	 *  For workbooks that follow the template, The full committee name is obtained if no abbreviated name is available
	 *  Finally the first sheet of the source workbook is cloned to a new destination sheet in the master budget
	 *
	 * @param color the tab color of the committeeSheet
	 * @param p the Portfolio to add the new Committee to
	 * @param committeeFile the xlsx file for this Committee
	 */
	static void createCommitteeBudget(IndexedColors color, Portfolio p, File committeeFile) {

		try (XSSFWorkbook committeeWB = new XSSFWorkbook(committeeFile)) {
			XSSFSheet committeeSheet = committeeWB.getSheetAt(0);

			XSSFName name = committeeWB.getName("NAME");
			CellReference nameRef = new CellReference(name.getRefersToFormula());
			XSSFRow nameRow = committeeSheet.getRow(nameRef.getRow());
			XSSFCell nameCell = nameRow.getCell(nameRef.getCol());
			String sheetName = nameCell.getStringCellValue();

			// Get full name if abbreviated unavailable in Template
			if (sheetName == null || sheetName.trim().length() == 0) {
				nameRow = committeeSheet.getRow(nameRef.getRow() - 1);
				nameCell = nameRow.getCell(nameRef.getCol());
				sheetName = nameCell.getStringCellValue();
			}

			// Source sheet name changed so that REV/EXP references have correct sheet name
			// for destination sheet
			committeeWB.setSheetName(0, sheetName);

			XSSFName revRef = committeeWB.getName("REV");
			XSSFName expRef = committeeWB.getName("EXP");

			new CommitteeBudget(sheetName, revRef.getRefersToFormula(), expRef.getRefersToFormula(), p);

			XSSFSheet bSheet = BudgetBuilder.budget.getWb().createSheet(sheetName);
			Cloner.cloneSheet(committeeSheet, bSheet);

			bSheet.setTabColor(color.getIndex());

			buildTask.updateBuildMessage("\t -" + sheetName);
			buildTask.updateBuildProgress(++currentProgress, totalProgress);
		} catch (IOException | InvalidFormatException e) {
			buildTask.updateBuildMessage(e.toString());
		}
	}
}
