package ca.mcgilleus.budgetbuilder.util;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 * Set of utility methods for cloning sheets, rows and cells even if the source and destination workbooks differ
 * Based off of code by Leonid Vvsochvn
 * @see <a href="http://jxls.cvs.sourceforge.net/jxls/jxls/src/java/org/jxls/util/Util.java?revision=1.8&view=markup"/>
 * @author Kareem Halabi
 */
public final class Cloner {
	
	private static List<CellRangeAddress> remainingRegions;
	
	public static void cloneSheet(XSSFSheet srcSheet, XSSFSheet destSheet) {
		remainingRegions = new ArrayList<>(srcSheet.getMergedRegions());
        int maxColumnNum = 0;
        for(int i = srcSheet.getFirstRowNum(); i <= srcSheet.getLastRowNum(); i++){
            XSSFRow srcRow = srcSheet.getRow( i );
            XSSFRow destRow = destSheet.createRow( i );
            if( srcRow != null ){
                cloneRow(srcRow, destRow);
                if( srcRow.getLastCellNum() > maxColumnNum ){
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for(int i = 0; i <= maxColumnNum; i++){
            destSheet.setColumnWidth( i, srcSheet.getColumnWidth( i ) );
        }
    }
	
	public static void cloneRow (XSSFRow srcRow, XSSFRow destRow) {
		
		destRow.setHeight( srcRow.getHeight() );

		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
			XSSFCell srcCell = srcRow.getCell(j);
			XSSFCell destCell = destRow.getCell(j);
			
			if(srcCell != null) {
				if(destCell == null) {
					destCell = destRow.createCell(j);
				}
				cloneCell(srcCell,destCell,true);
				
				//Merged regions have to be dealt with separately and prevented from being overlapped
				CellRangeAddress mergedRegion = getMergedRegion(srcCell, remainingRegions);
				if(mergedRegion != null) {
					destRow.getSheet().addMergedRegion(mergedRegion);
				}
			}
		}
	}

	public static CellRangeAddress getMergedRegion(XSSFCell srcCell, List<CellRangeAddress> regions) {

		for(CellRangeAddress region : regions) {
			if(region.isInRange(srcCell.getRowIndex(), srcCell.getColumnIndex())) {
				regions.remove(region);
				return region;
			}
				
		}
		return null;
	}

	public static void cloneCell(XSSFCell srcCell, XSSFCell destCell, boolean copyStyle) {
        if( copyStyle ){
        	CellStyle srcStyle = srcCell.getCellStyle();
        	
        	XSSFWorkbook destWB = destCell.getSheet().getWorkbook();
        	CellStyle destStyle = destWB.createCellStyle();
        	destStyle.cloneStyleFrom(srcStyle);
            destCell.setCellStyle(destStyle);
        }
        switch (srcCell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                destCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                destCell.setCellErrorValue(srcCell.getErrorCellValue());
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
            	//Fix for cells containing external references
				if(srcCell.getCellFormula().contains("!") &&
						!srcCell.getCellFormula().contains("#REF!") &&
						!srcCell.getSheet().getWorkbook().equals(destCell.getSheet().getWorkbook())
				)
					destCell.setCellValue(srcCell.getNumericCellValue());
				else
					destCell.setCellFormula(srcCell.getCellFormula());
                break;
            default:
                break;
        }
    }
}
