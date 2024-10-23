package task8;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sheet1 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		//get into the workbook
		
		XSSFWorkbook book = new XSSFWorkbook();
		
		//create the sheet
		
		XSSFSheet sheet = book.createSheet("Details");
		
		//store the details -> Name (String)  Age(int) city(String)
		
		Object[][] data = {
				
				{"Name","Age","Email"},
				{"John doe",30,"john@test.com"},
				{"Jane doe",28,"john@test.com"},
				{"Bob Smith",35,"jacky@example.com"},
				{"Swapnil",37,"swapnil@example.com"}
				
		};
		
		// Put the data into the sheet
		
		int rowCount = 0;
		
		// for each to get into each row
		
		for(Object[] row1 : data ) {
			
			XSSFRow row = sheet.createRow(rowCount++);
			
			int columnCount =0;
			
			// for each to get the columns
			
			for(Object col : row1) {
				
			XSSFCell cell = row.createCell(columnCount++);
			
			// Checking the type of data and making the entry
			if(col instanceof String) {
				
				cell.setCellValue((String)col);
				
			} else if (col instanceof Integer) {
				
				cell.setCellValue((Integer)col);
			}
			}
		}
		try {
			FileOutputStream output = new FileOutputStream("C:\\Users\\siddh\\eclipse-workspace\\ExcelFileOperations\\src\\main\\java\\task8\\Sheet1.xlxs");
			book.write(output);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		book.close();
	}

}
