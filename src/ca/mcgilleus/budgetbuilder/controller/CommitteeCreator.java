package ca.mcgilleus.budgetbuilder.controller;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;
import ca.mcgilleus.budgetbuilder.util.Cloner;

public class CommitteeCreator {

	public static void createCommitteeBudget(IndexedColors color, Portfolio p, File pF)
			throws IOException, InvalidFormatException {

		XSSFWorkbook pWB = (XSSFWorkbook) WorkbookFactory.create(pF);
		XSSFSheet pSheet = pWB.getSheetAt(0);

		Name name = pWB.getName("COMM_NAME");
		CellReference nameRef = new CellReference(name.getRefersToFormula());
		XSSFRow nameRow = pSheet.getRow(nameRef.getRow());
		XSSFCell nameCell = nameRow.getCell(nameRef.getCol());
		String sheetName = nameCell.getStringCellValue();

		// Get full name if abbreviated unavailable
		if (sheetName == null || sheetName.trim().length() == 0) {
			nameRow = pSheet.getRow(nameRef.getRow() - 1);
			nameCell = nameRow.getCell(nameRef.getCol());
			sheetName = nameCell.getStringCellValue();
		}

		// So that AMT reference has correct sheet name
		pWB.setSheetName(0, sheetName);

		Name amt = pWB.getName("AMT");

		CommitteeBudget committeeBudget = new CommitteeBudget(sheetName, amt.getRefersToFormula(), p);

		XSSFSheet bSheet = EUSBudgetBuilder.getWorkbook().createSheet(sheetName);
		Cloner.cloneSheet(pSheet, bSheet);

		bSheet.setTabColor(color.getIndex());
		System.out.println("\t -" + sheetName);
	}

}
