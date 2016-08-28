package ca.mcgilleus.budgetbuilder.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;

import static ca.mcgilleus.budgetbuilder.service.BudgetBuilder.budget;

/**
 * Contains all fields and methods relevant to styling cells
 * @author Kareem Halabi
 */
public final class Styles {
	
	public static XSSFCellStyle HEADER_STYLE;
	public static XSSFCellStyle PORTFOLIO_LABEL_STYLE;
	public static XSSFCellStyle COMMITTEE_LABEL_STYLE;
	public static XSSFCellStyle CURRENCY_CELL_STYLE;
	public static XSSFCellStyle AMT_CURRENCY_CELL_STYLE;
	public static XSSFCellStyle TOTAL_LABEL_STYLE;
	public static XSSFCellStyle TOTAL_CELL_STYLE;

	private static int currentColor = 0;
	private static IndexedColors[] colors = {
		IndexedColors.MAROON,
		IndexedColors.LIGHT_ORANGE,
		IndexedColors.YELLOW,
		IndexedColors.GREEN,
		IndexedColors.LIGHT_BLUE,
		IndexedColors.LAVENDER,
		IndexedColors.GREY_40_PERCENT
	};

	//initialize styles
	public static void initStyles() {
		HEADER_STYLE = budget.getWb().createCellStyle();
		XSSFFont headerFont= budget.getWb().createFont();
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

		PORTFOLIO_LABEL_STYLE = budget.getWb().createCellStyle();
		XSSFFont basicFont= budget.getWb().createFont();
		basicFont.setFontHeightInPoints((short)10);
		basicFont.setFontName("Arial");
		basicFont.setColor(IndexedColors.WHITE.getIndex());
		basicFont.setBold(false);
		basicFont.setItalic(false);

		PORTFOLIO_LABEL_STYLE.setFont(basicFont);

		PORTFOLIO_LABEL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		PORTFOLIO_LABEL_STYLE.setAlignment(CellStyle.ALIGN_CENTER);

		PORTFOLIO_LABEL_STYLE.setFillForegroundColor(IndexedColors.GREY_80_PERCENT.getIndex());
		PORTFOLIO_LABEL_STYLE.setFillPattern(CellStyle.SOLID_FOREGROUND);

		//----------------------------------------------------------------

		COMMITTEE_LABEL_STYLE = budget.getWb().createCellStyle();

		XSSFFont committeeFont = budget.getWb().createFont();
		committeeFont.setFontHeightInPoints((short)10);
		committeeFont.setFontName("Arial");
		committeeFont.setColor(IndexedColors.BLACK.getIndex());
		committeeFont.setBold(false);
		committeeFont.setItalic(false);

		COMMITTEE_LABEL_STYLE.setFont(committeeFont);

		COMMITTEE_LABEL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		COMMITTEE_LABEL_STYLE.setAlignment(CellStyle.ALIGN_LEFT);

		COMMITTEE_LABEL_STYLE.setBorderRight(CellStyle.BORDER_THIN);
		COMMITTEE_LABEL_STYLE.setRightBorderColor(IndexedColors.BLACK.getIndex());


		//----------------------------------------------------------------

		CURRENCY_CELL_STYLE = budget.getWb().createCellStyle();

	    CURRENCY_CELL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		CURRENCY_CELL_STYLE.setAlignment(CellStyle.ALIGN_RIGHT);

		CURRENCY_CELL_STYLE.setBorderRight(CellStyle.BORDER_THIN);
		CURRENCY_CELL_STYLE.setRightBorderColor(IndexedColors.BLACK.getIndex());

		XSSFDataFormat cF = budget.getWb().createDataFormat();
		//Positive values: "$ 1,234,567.89"
		// Negative values (in red): "-$ 1,234,567.89"
		// Zero: "$  -  "
		short formatIndex = cF.getFormat("$ #,##0.00;[Red]-$ #,##0.00;_-$??\"-\"??_-;_-@_-");

		CURRENCY_CELL_STYLE.setDataFormat(formatIndex);

		//----------------------------------------------------------------

		AMT_CURRENCY_CELL_STYLE = budget.getWb().createCellStyle();

		AMT_CURRENCY_CELL_STYLE.cloneStyleFrom(CURRENCY_CELL_STYLE);

		XSSFColor grey15Percent = new XSSFColor(new byte[] {(byte) 217,(byte) 217,(byte) 217});

		AMT_CURRENCY_CELL_STYLE.setFillForegroundColor(grey15Percent);
		AMT_CURRENCY_CELL_STYLE.setFillPattern(CellStyle.SOLID_FOREGROUND);

		//----------------------------------------------------------------

		TOTAL_LABEL_STYLE = budget.getWb().createCellStyle();
		XSSFFont totalFont= budget.getWb().createFont();
	    totalFont.setFontHeightInPoints((short)10);
	    totalFont.setFontName("Arial");
	    totalFont.setColor(IndexedColors.BLACK.getIndex());
	    totalFont.setBold(true);
	    totalFont.setItalic(false);

		TOTAL_LABEL_STYLE.setFont(totalFont);

	    TOTAL_LABEL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		TOTAL_LABEL_STYLE.setAlignment(CellStyle.ALIGN_LEFT);

		//----------------------------------------------------------------

		TOTAL_CELL_STYLE = budget.getWb().createCellStyle();

	    TOTAL_CELL_STYLE.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		TOTAL_CELL_STYLE.setAlignment(CellStyle.ALIGN_RIGHT);

		TOTAL_CELL_STYLE.setBorderTop(CellStyle.BORDER_DOUBLE);
		TOTAL_CELL_STYLE.setTopBorderColor(IndexedColors.BLACK.getIndex());

		TOTAL_CELL_STYLE.setDataFormat((short) 8); // ($#,##0.00_);[Red]($#,##0.00)
	}

	public static IndexedColors popTabColor() {
		return colors[(currentColor++)%colors.length];
	}

	public static XSSFCellStyle getPortfolioLabelStyle(IndexedColors color) {
		XSSFCellStyle customPortfolioStyle = budget.getWb().createCellStyle();
		customPortfolioStyle.cloneStyleFrom(Styles.PORTFOLIO_LABEL_STYLE);
		XSSFFont customPortfolioFont = budget.getWb().createFont();
		customPortfolioFont.setFontHeightInPoints((short)10);
		customPortfolioFont.setFontName("Arial");
		customPortfolioFont.setColor(color.getIndex());
		customPortfolioFont.setBold(false);
		customPortfolioFont.setItalic(false);
		customPortfolioStyle.setFont(customPortfolioFont);

		return customPortfolioStyle;
	}
}
