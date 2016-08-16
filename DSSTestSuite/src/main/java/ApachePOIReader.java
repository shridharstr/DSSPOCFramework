import java.awt.List;
import java.io.*;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.dss.test.properties.*;

public class ApachePOIReader {

	public static XSSFCell cell;
	public static XSSFCell Cell;
	public static XSSFWorkbook wbook;
	public static XSSFSheet Sheet;
	public static FileInputStream inputStream;
	
	 public static void getCellData() throws IOException
		{
		    Workbook wb = new XSSFWorkbook(inputStream);
			org.apache.poi.ss.usermodel.Sheet Firstsheet = wb.getSheet("OS");
			int totalMergedRegions = Firstsheet.getNumMergedRegions();
			for(int i=0; i< totalMergedRegions; i++)
			{
				CellRangeAddress ca = Firstsheet.getMergedRegion(i);
				int mergedRow = ca.getFirstRow();
			}	
			
		
            
			
	    }
	 
	public static void main(String[] args) throws Exception
	{
		String excelFile = 	DSSProperties.filePath;	
		inputStream = new FileInputStream(new File(excelFile));
		getCellData();
		
	}

}
