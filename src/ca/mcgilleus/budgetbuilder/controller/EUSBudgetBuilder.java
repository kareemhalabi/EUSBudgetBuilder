package ca.mcgilleus.budgetbuilder.controller;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.Queue;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.model.CommitteeBudget;
import ca.mcgilleus.budgetbuilder.model.EUSBudget;
import ca.mcgilleus.budgetbuilder.model.Portfolio;

/**
 * @author Kareem Halabi
 */

public class EUSBudgetBuilder {
	
	private static XSSFWorkbook wb;
	
	public static XSSFWorkbook getWorkbook() {
		return wb;
	}

	public static void main(String[] args) {
		
		File root = new File(System.getProperty("user.dir") + File.separator + "W2017 Budget");
		EUSBudget budget = createBudget();
		wb = budget.getWorkbook();
		
		wb.createSheet(budget.getBudgetYear() + " Budget");
		
		Queue<IndexedColors> availableTabColors = new LinkedList<IndexedColors>();
		createColorQueue(availableTabColors);
		
		PortfolioCreator.setAllColors(availableTabColors);
		
		for(File f : root.listFiles()) {
			PortfolioCreator.createPortfolio(f, budget);
		}
System.out.println("Compiling EUS Budget Overview");		
		createEUSBudgetOverview(budget);
		
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(root.getName() + ".xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
			
			System.out.println("Openning: " + root.getName() + ".xlsx");
			Desktop.getDesktop().open(new File(root.getName() + ".xlsx"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	private static void createEUSBudgetOverview(EUSBudget budget) {
		
		XSSFSheet overviewSheet = wb.getSheet(budget.getBudgetYear() + " Budget");
		XSSFRow header = overviewSheet.createRow(0);
		
		XSSFCell title = header.createCell(0, Cell.CELL_TYPE_STRING);
		title.setCellValue("Portfolio");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		title.setCellValue("Committee");
		
		title = header.createCell(header.getLastCellNum(), Cell.CELL_TYPE_STRING);
		
		title.setCellValue(budget.getBudgetYear() + " Budget");
		
		for(int i = 0; i < budget.getPortfolios().size(); i++) {
			Portfolio p = budget.getPortfolio(i);
			for(int j = 0; j < p.getCommitteeBudgets().size(); j++) {
				CommitteeBudget committee = p.getCommitteeBudget(j);
				XSSFRow row = overviewSheet.createRow(overviewSheet.getLastRowNum()+1);
				
				XSSFCell portfolio = row.createCell(0, Cell.CELL_TYPE_STRING);
				portfolio.setCellValue(p.getName());
				
				XSSFCell committeeName = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_STRING);
				committeeName.setCellValue(committee.getName());
				
				XSSFCell committeeAmt = row.createCell(row.getLastCellNum(), Cell.CELL_TYPE_FORMULA);
				committeeAmt.setCellFormula(committee.getAmtRequestedRef());	
			}
		}
	}

	public static void createColorQueue(Queue<IndexedColors> colors) {
		colors.add(IndexedColors.MAROON);
		colors.add(IndexedColors.LIGHT_ORANGE);
		colors.add(IndexedColors.YELLOW);
		colors.add(IndexedColors.GREEN);
		colors.add(IndexedColors.BLUE);
		colors.add(IndexedColors.PLUM);
		colors.add(IndexedColors.GREY_40_PERCENT);
		
	}


	public static EUSBudget createBudget() {
		Calendar c = Calendar.getInstance();
		Date budgetYear;
		if(c.get(Calendar.MONTH) <= Calendar.FEBRUARY)
			c.add(Calendar.YEAR, -1);
		
		budgetYear = c.getTime();
		
		return new EUSBudget(budgetYear);
	}
	
}
