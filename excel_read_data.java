package Excel_Data;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class excel_read_data {

	
	public static void main(String[] args) throws Exception {
		
		String s = System.getProperty("user.dir")+"\\Excel\\data.xlsx";
		FileInputStream fis = new FileInputStream(s);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int Total_Sheets = workbook.getNumberOfSheets();
		
		for(int i =0;i<Total_Sheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("one")) {
				
				// Identify the Sheet
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				//Identify  the each row
				Iterator<Row> rows = sheet.rowIterator();
				Row row = rows.next();
				Iterator<Cell> cell = row.cellIterator();
				int k =0;
				int column=0;
				while(cell.hasNext()) {
					Cell value = cell.next();
					if(value.getStringCellValue().equalsIgnoreCase("testcases")) {
						column=k;
					}
					k++;
				}
				System.out.println("Column Number    "+column);
				
				while(rows.hasNext()) {
					row = rows.next();
					if(row.getCell(column).getStringCellValue().equalsIgnoreCase("purchase")) {
						cell = row.cellIterator();
						while(cell.hasNext()) {
							System.out.println(cell.next().getStringCellValue());
						}
					}
				}
				
			}
		}
		
	}
}
