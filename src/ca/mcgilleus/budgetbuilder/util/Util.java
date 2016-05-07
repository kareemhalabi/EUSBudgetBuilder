package ca.mcgilleus.budgetbuilder.util;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Based off of code by
 * @author Leonid Vvsochvn
 * {@link http://jxls.cvs.sourceforge.net/jxls/jxls/src/java/org/jxls/util/Util.java?revision=1.8&view=markup}
 */

public final class Util {
	
	private static List<CellRangeAddress> remainningRegions;
	
	public static void cloneSheet(XSSFSheet srcSheet, XSSFSheet destSheet) {
		remainningRegions = new ArrayList<CellRangeAddress>(srcSheet.getMergedRegions());
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
				CellRangeAddress mergedRegion =  getMergedRegion(srcCell, remainningRegions);
				if(mergedRegion != null) {
					CellRangeAddress newMergedRegion = new CellRangeAddress(
							mergedRegion.getFirstRow(), mergedRegion.getLastRow(),
							mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
					
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
        	XSSFWorkbook srcWB = srcCell.getSheet().getWorkbook();
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
            	destCell.setCellFormula(srcCell.getCellFormula());
                break;
            default:
                break;
        }
    }
	
}
