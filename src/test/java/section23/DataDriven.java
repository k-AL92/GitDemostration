package section23;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;


public class DataDriven {
	
	public ArrayList<String> getData(String testcaseName) throws IOException {
		ArrayList<String> data = new ArrayList<String>();
		FileInputStream file = new FileInputStream("C:\\Users\\karina.andujo\\Documents\\TestingExcel.xlsx");
		XSSFWorkbook book = new XSSFWorkbook(file);
		int noSheets = book.getNumberOfSheets();
		for(int i=0;i<noSheets;i++) {
			if(book.getSheetName(i).equalsIgnoreCase("testData")) {
				XSSFSheet sheet = book.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> cell = firstRow.cellIterator();
				int j= 0;
				int column=0;
				while(cell.hasNext()) {
					Cell cellValue = cell.next();
					if(cellValue.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column = j;
						
					}
					j++;
				}
				while (rows.hasNext()) {
					Row row =rows.next();
					if(row.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						Iterator <Cell> c = row.cellIterator();
						while(c.hasNext()) {
							Cell celda = c.next();
							if(celda.getCellType()==CellType.STRING) {
								data.add(celda.getStringCellValue());
							}else {
								data.add(NumberToTextConverter.toText(celda.getNumericCellValue()));								;
							}
							
						}
					}
				}
			}
			
		}
		return data;
	}

}
