package ca.mcgilleus.budgetbuilder.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFFont;

import ca.mcgilleus.budgetbuilder.controller.EUSBudgetBuilder;

public final class Styles {
	
	public static final CellStyle HEADER_STYLE;
	public static final CellStyle PORTFOLIO_LABEL_STYLE;
	public static final CellStyle COMMITTEE_LABEL_STYLE;
	public static final CellStyle CURRENCY_CELL_STYLE;
	public static final CellStyle TOTAL_LABEL_STYLE;
	public static final CellStyle TOTAL_CELL_STYLE;

	//initialize styles
	static {
		HEADER_STYLE = EUSBudgetBuilder.getWorkbook().createCellStyle();
		XSSFFont headerFont= EUSBudgetBuilder.getWorkbook().createFont();
	    headerFont.setFontHeightInPoints((short)11);
	    headerFont.setFontName("Arial");
	    headerFont.setColor(IndexedColors.WHITE.getIndex());
	    headerFont.setBold(true);
	    headerFont.setItalic(false);
	
		HEADER_STYLE.setFont(headerFont);
		
		HEADER_STYLE.setWrapText(true);
	    
		HEADER_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		HEADER_STYLE.setAlignment(CellStyle.ALIGN_CENTER);

		HEADER_STYLE.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		HEADER_STYLE.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		//----------------------------------------------------------------
		
		PORTFOLIO_LABEL_STYLE = EUSBudgetBuilder.getWorkbook().createCellStyle();
		XSSFFont basicFont= EUSBudgetBuilder.getWorkbook().createFont();
		basicFont.setFontHeightInPoints((short)10);
		basicFont.setFontName("Arial");
		basicFont.setColor(IndexedColors.BLACK.getIndex());
		basicFont.setBold(false);
		basicFont.setItalic(false);
		
		PORTFOLIO_LABEL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		PORTFOLIO_LABEL_STYLE.setAlignment(CellStyle.ALIGN_CENTER);
		
		PORTFOLIO_LABEL_STYLE.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
		PORTFOLIO_LABEL_STYLE.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		//----------------------------------------------------------------
		
		COMMITTEE_LABEL_STYLE = EUSBudgetBuilder.getWorkbook().createCellStyle();
		
		COMMITTEE_LABEL_STYLE.setFont(basicFont);
		
		COMMITTEE_LABEL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		COMMITTEE_LABEL_STYLE.setAlignment(CellStyle.ALIGN_LEFT);
		
		COMMITTEE_LABEL_STYLE.setBorderRight(CellStyle.BORDER_THIN);
		COMMITTEE_LABEL_STYLE.setRightBorderColor(IndexedColors.BLACK.getIndex());
		
		
		//----------------------------------------------------------------
		
		CURRENCY_CELL_STYLE = EUSBudgetBuilder.getWorkbook().createCellStyle();
		
	    CURRENCY_CELL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		CURRENCY_CELL_STYLE.setAlignment(CellStyle.ALIGN_RIGHT);
		
		CURRENCY_CELL_STYLE.setBorderRight(CellStyle.BORDER_THIN);
		CURRENCY_CELL_STYLE.setRightBorderColor(IndexedColors.BLACK.getIndex());
		
		CURRENCY_CELL_STYLE.setDataFormat((short) 8); // ($#,##0.00_);[Red]($#,##0.00)
		
		//----------------------------------------------------------------
		
		TOTAL_LABEL_STYLE = EUSBudgetBuilder.getWorkbook().createCellStyle();
		XSSFFont totalFont= EUSBudgetBuilder.getWorkbook().createFont();
	    totalFont.setFontHeightInPoints((short)10);
	    totalFont.setFontName("Arial");
	    totalFont.setColor(IndexedColors.BLACK.getIndex());
	    totalFont.setBold(true);
	    totalFont.setItalic(false);
		
		TOTAL_LABEL_STYLE.setFont(totalFont);
		
	    TOTAL_LABEL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		TOTAL_LABEL_STYLE.setAlignment(CellStyle.ALIGN_LEFT);
		
		//----------------------------------------------------------------
		
		TOTAL_CELL_STYLE = EUSBudgetBuilder.getWorkbook().createCellStyle();
		
	    TOTAL_CELL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		TOTAL_CELL_STYLE.setAlignment(CellStyle.ALIGN_RIGHT);
		
		TOTAL_CELL_STYLE.setBorderTop(CellStyle.BORDER_DOUBLE);
		TOTAL_CELL_STYLE.setTopBorderColor(IndexedColors.BLACK.getIndex());
		
		TOTAL_CELL_STYLE.setDataFormat((short) 8); // ($#,##0.00_);[Red]($#,##0.00)
	}

}
