package excelloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel {
	public static void main(String[] args) throws IOException {
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet = workbook.createSheet("Employee Info");
	Object empdata[][] = {
					{"EmpID","Name","Job"},
					{"101","Daipayan","Software Engineer"},
					{"102","Virat","Cricketer"},
					{"103","Amitabh","Actor"}
			};
//	Using Nomral Form Loop
	int rows = empdata.length;
	int cols = empdata[0].length;
	for(int r = 0;r<rows;r++) {
		XSSFRow cur_row =  sheet.createRow(r);
		for(int c = 0;c<cols;c++) {
			XSSFCell cur_cell = cur_row.createCell(c);
			Object value  = empdata[r][c];
			if(value instanceof String) {
				cur_cell.setCellValue((String)value);
			}
			if(value instanceof Integer) {
				cur_cell.setCellValue((Integer)value);
			}
			if(value instanceof Boolean) {
				cur_cell.setCellValue((Boolean)value);
			}
		}
	}
	String filePath = ".\\DataFiles\\employee.xlsx";
	FileOutputStream outStream = new FileOutputStream(filePath);
	workbook.write(outStream);
	outStream.close();
	System.out.println("File Writen Success");
	}
}
