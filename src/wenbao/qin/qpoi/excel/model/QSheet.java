package wenbao.qin.qpoi.excel.model;

/***
 * 
 * Sheet Object
 *
 */
public class QSheet {

	/**sheet name **/
	private String sheetName ;
	
	/**sheet index **/
	private String sheetIndex ;
	
	/**sheet col count **/
	private String sheetColCount ;
	
	/**sheet row count **/
	private String sheetRowCount ;

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public String getSheetIndex() {
		return sheetIndex;
	}

	public void setSheetIndex(String sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	public String getSheetColCount() {
		return sheetColCount;
	}

	public void setSheetColCount(String sheetColCount) {
		this.sheetColCount = sheetColCount;
	}

	public String getSheetRowCount() {
		return sheetRowCount;
	}

	public void setSheetRowCount(String sheetRowCount) {
		this.sheetRowCount = sheetRowCount;
	}
}
