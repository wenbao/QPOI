package wenbao.qin.qpoi.excel.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Random;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

import wenbao.qin.qpoi.excel.model.QExcel;
import wenbao.qin.qpoi.excel.util.QUtil;

/*
 * Read Excel Object
 */
public class QReadExcelImpl implements QReadExcel{
	
	Logger logger = Logger.getLogger(this.getClass().getName()) ;
	
	private QExcel excel ;
	
	public QReadExcelImpl( QExcel excel ) {
		this.excel = excel ;
	} 
	
	/*
	 * read the excel
	 */
	public void read() throws IOException
	{
		 
	}
	
	public HSSFWorkbook getWorkBook() 
	{
		File f = new File( excel.getExcelPath() ) ; 
		InputStream input;
		HSSFWorkbook workBook = null  ;
		try 
		{
			input = new FileInputStream( f ) ;
			try 
			{
				workBook = new HSSFWorkbook( input ) ;
			} 
			catch (IOException e) 
			{
				e.printStackTrace();
			}
			return workBook ; 
		} 
		catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		} 
		return null ;
	}
	
	public HSSFSheet getSheet ( int sheetIndex)  
	{
		HSSFWorkbook workBook = getWorkBook() ;
		HSSFSheet sheet = workBook.getSheetAt( sheetIndex ) ;
		return sheet ;
	}
	
	public List getAllSheet()  
	{
		List list = new LinkedList() ;
		HSSFWorkbook workBook = getWorkBook() ;
		int sheetCount = workBook.getNumberOfSheets() ;
		HSSFSheet sheet = null ;
		for( int i=0; i < sheetCount; i++ )
		{
			sheet = workBook.getSheetAt( i) ;
			list.add(sheet) ;
		}
		return list ;
	}
	
	
	public static void main(String[] args) 
	{
//		String excelPath  = "D:/tmp/数据中国数据表总表@@渝中区.xls" ;
		String excelPath  = "C:/Users/qinwenbao/Desktop/五大功能区加载数据V1.0.xls" ;
		int sheetIndex = 0 ;
		QExcel excel = new QExcel(excelPath) ;
		QReadExcel reader = new QReadExcelImpl( excel ) ; 
		reader.scaneSheet(reader.getSheet( sheetIndex )) ;
		
		
//		System.out.println("" + excel.getExcelName() ) ;
		
	}
	
	@Override
	public void scaneSheet( HSSFSheet sheet )
	{
		int rowNum = sheet.getLastRowNum() + 1 ;
		int colNum = sheet.getRow(0).getLastCellNum() ;
		
		HSSFRow row ;
		HSSFCell cell ;
		
		
		String time = sheet.getSheetName() ; 
		List themeList = new ArrayList() ; 
//		List themeList = new LinkedList() ; 
//		List themeList = new LinkedList() ; 
//		List themeList = new LinkedList() ; 
		
		String navStr = "行号,id,主题,类别,指标,指标涵义," + time ;
		System.out.println(navStr);
		
		String tmp  ;
		int curIndex = -1 ;
		for( int i = 0; i < colNum; i++ ) 
		{
			tmp = "" ;
			for( int j = 0; j < rowNum; j++ )
			{
				row = sheet.getRow( j ) ;
				cell = row.getCell(i) ;
				if( i == 0 ) //地区
				{
					String cellValue = QUtil.getCellValue(cell) ;
					if ( !"".equals(cellValue) )
					{
						themeList.add( cellValue ) ;
					}
				}
				else  
				{
					String cellValue = QUtil.getCellValue(cell) ;
					
					if( j <= 3 )
					{
						tmp += ( cellValue + "," ) ; 
						
					}
					else
					{
						curIndex = tmp.indexOf(",") ;
						System.out.println(j+i+","+ tmp.substring(0, curIndex) +","+ themeList.get(j-3-1)
														+ tmp.substring(curIndex, tmp.length()) +"" + cellValue
														+ ",,"
														);
					}
				}
			}
		}
		
	}
	
	
	
	
		
}
