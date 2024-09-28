package utils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("StudentData1");
		
		 Object studata [][] = { {"Name","Age","Email"},{"John",25,"john@test.com"},
				 {"Smith",35,"smith@test.com"},{"Nick",20,"nick@test.com"}
		 
		 
		 };
		 int rows = studata.length;
		 int cols = studata[0].length;
		 
		 System.out.println(rows);//4
		 System.out.println(cols);//3
		 
		 for(int r = 0; r < rows; r++) {
			 XSSFRow row = sheet.createRow(r);
			 
			 for(int c = 0; c < cols;c++) {
				 
			 XSSFCell cell  = row.createCell(c);
			 Object value = studata[r][c];
			 if(value instanceof String)
				 cell.setCellValue((String)value);
			 if(value instanceof Integer)
				 cell.setCellValue((Integer)value);
			 if(value instanceof Boolean)
				 cell.setCellValue((Boolean)value);
			 
				 
			 }
		 }
		 String filePath = ".\\Excellstudentdata\\Studentdata1.xlsx";
		 FileOutputStream fos = new FileOutputStream(filePath);
		 workbook.write(fos);
		 fos.close();
		 System.out.println("Student Data file Created Successfully!!");

	}

}
