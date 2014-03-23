package wenbao.qin.qpoi.excel.util;

import java.text.DecimalFormat;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;

public class QUtil {
	
	static Logger logger = Logger.getLogger( QUtil.class.getName() ) ;
	
	/**
	 * 	fit differ cell type ,return string 
	 */
	public static String getCellValue(HSSFCell cell) {
		
		if( null == cell)
		{
//			logger.log( Level.WARNING,"cell is null." ) ;
			return "";
		}
		String cellValue = "";
		DecimalFormat df = new DecimalFormat("#");
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			cellValue = cell.getRichStringCellValue().getString().trim();
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
//			cellValue = df.format(cell.getNumericCellValue()).toString();
			cellValue =  String.valueOf( cell.getNumericCellValue() ) ;
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			cellValue = cell.getCellFormula();
			break;
		default:
			cellValue = "";
		}
		return cellValue;
	}
	
	/**
	 *   get the cell value of the display  
	 */
	public static String getShowStringValue( HSSFCell cell )
	{
		HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
		String cellFormatted = dataFormatter.formatCellValue( cell );
		dataFormatter = null ;
		return cellFormatted ;
	}

}
