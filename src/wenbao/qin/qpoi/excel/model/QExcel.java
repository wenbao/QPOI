package wenbao.qin.qpoi.excel.model;

import java.util.logging.Level;
import java.util.logging.Logger;

import wenbao.qin.util.file.QFile;
import wenbao.qin.util.file.impl.QFileImpl;

public class QExcel {
	
	Logger logger = Logger.getLogger(this.getClass().getName()) ;
	
	public QExcel(String excelPath) {
		this.excelPath = excelPath ;
		validate() ;
		getExcelNameByExcelPath() ;
	}
	
	/*
	 *	the excel path  
	 */
	private String excelPath ;

	/*
	 *	the excel name  
	 */
	private String excelName ;
	
	/*
	 *	the excel type
	 *	xls or xlsx  
	 */
	private String excelType ;
	
	/*
	 *	the excel version
	 *	2003/2007/2010/2013
	 *  
	 */
	private String excelVersion ;

	
	public String getExcelPath() {
		return excelPath;
	}

	public void setExcelPath(String excelPath) {
		this.excelPath = excelPath;
	}

	public String getExcelName() {
		return excelName;
	}

	public void setExcelName(String excelName) {
		this.excelName = excelName;
	}

	public String getExcelType() {
		return excelType;
	}

	public void setExcelType(String excelType) {
		this.excelType = excelType;
	}

	public String getExcelVersion() {
		return excelVersion;
	}

	public void setExcelVersion(String excelVersion) {
		this.excelVersion = excelVersion;
	}
	
	private void validate( )
	{
		QFile f = new QFileImpl( this.excelPath ) ;
		f.checkFilePath() ;
	}
	
	private void getExcelNameByExcelPath()
	{
		int index = this.excelPath.indexOf("/") ;
		if( index != -1 )
		{
			int lastIndex = excelPath.lastIndexOf( "/" ) + 1 ; 
			this.excelName = this.excelPath.substring(lastIndex, excelPath.length()) ;
		}
	}
	
}
