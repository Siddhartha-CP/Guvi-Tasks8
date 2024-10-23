package day16;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		//open the workbook
		
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\siddh\\eclipse-workspace\\ExcelFileOperations\\src\\main\\java\\day16\\studentDetails.xlxs");
		
		// get into the sheet
		
		XSSFSheet sheet = book.getSheet("details");
		
		//get the no of rows
		
		int rowCount = sheet.getLastRowNum();
		
		//get the no.of columns
		
		int columnCount = sheet.getRow(0).getLastCellNum();
		
		// create 2D array
		
		String[][] data = new String[rowCount][columnCount];
		
		for(int i =1 ; i<= rowCount; i++ ) {
			
			XSSFRow row = sheet.getRow(i);
			
			// get into the columns
			
			for(int j=0; j<columnCount; j++) {
				XSSFCell cell = row.getCell(j);
				
				// read/get the value
				
				System.out.println(cell.getStringCellValue()); 
				
				//to store in array
				data[i-1][j] = cell.getStringCellValue();
			}
			System.out.println();
		}
		for (String[]row : data) {
			for(String x : row) {
				System.out.println(x+" ");
			}
		}
		book.close();
	}

}
