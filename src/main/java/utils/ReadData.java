package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public static void main(String[] args) {
		getRowCount();
		//getCellDataString();
		getCellDataNumeric();
		}

	public static void getRowCount () {
		try {
			workbook = new XSSFWorkbook("C:\\Users\\sivah\\prinkyworkspace\\Excellproject\\Excellstudentdata\\Studentdata.xlsx");
		
		sheet = workbook.getSheet("Sheet1");
		int rawCount = sheet.getPhysicalNumberOfRows();
		System.out.println("Number of Rows: "+rawCount);
		
		
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	public static void getCellDataString() {
		try {
			workbook = new XSSFWorkbook("C:\\Users\\sivah\\prinkyworkspace\\Excellproject\\Excellstudentdata\\Studentdata.xlsx");
		
		sheet = workbook.getSheet("Sheet1");
		String cellData = sheet.getRow(1).getCell(1).getStringCellValue();
		
		System.out.println(cellData);
		
		
		} catch (IOException e) {
			e.printStackTrace();	
	}
	
}
	public static void getCellDataNumeric(){
		try {
			workbook = new XSSFWorkbook("C:\\Users\\sivah\\prinkyworkspace\\Excellproject\\Excellstudentdata\\Studentdata.xlsx");
		
		sheet = workbook.getSheet("Sheet1");
		double cellData = sheet.getRow(1).getCell(1).getNumericCellValue();
		
		System.out.println(cellData);
		
		
		} catch (IOException e) {
			e.printStackTrace();
}
	}
}

