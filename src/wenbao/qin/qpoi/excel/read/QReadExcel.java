package wenbao.qin.qpoi.excel.read;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Read Excel Object
 */
public interface QReadExcel {

	/**
	 *  get workbook
	 */
	public HSSFWorkbook getWorkBook()  ;
	
	/**
	 *	get all sheet object 
	 */
	public List getAllSheet()  ; 
	
	/**
	 *	get specify index sheet 
	 */
	public HSSFSheet getSheet( int sheetIndex)  ;
	
	/**
	 *  scane sheet   
	 */
	public void scaneSheet( HSSFSheet sheet ) ;
	
}
