package ca.mcgilleus.budgetbuilder.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.Queue;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ca.mcgilleus.budgetbuilder.model.EUSBudget;

/**
 * Kareem Halabi
 * 260 616 162
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
		
		Queue<IndexedColors> availableTabColors = new LinkedList<IndexedColors>();
		createColorQueue(availableTabColors);
		
		PortfolioCreator.setAllColors(availableTabColors);
		
		for(File f : root.listFiles()) {
			PortfolioCreator.createPortfolio(f, budget);
		}
		
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(root.getName() + ".xlsx");
			wb.write(fileOut);
			fileOut.close();
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	public static void createColorQueue(Queue<IndexedColors> colors) {
		colors.add(IndexedColors.MAROON);
		colors.add(IndexedColors.LIGHT_ORANGE);
		colors.add(IndexedColors.LIGHT_YELLOW);
		colors.add(IndexedColors.LIGHT_GREEN);
		colors.add(IndexedColors.LIGHT_BLUE);
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
