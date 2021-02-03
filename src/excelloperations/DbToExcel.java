package excelloperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DbToExcel {
	public static void main(String[] args) throws SQLException, IOException {
//		Connection 2 DataBase
		Connection con =  DriverManager.getConnection("jdbc:mysql://localhost:3306/ems","root","pass");
		System.out.println("Con Success");
//		Statement Query 1
		Statement stm = con.createStatement();
		ResultSet result =  stm.executeQuery("select * from employee");
//		Excell
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet =  workbook.createSheet("EmployeeData");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("emp_id");
		row.createCell(1).setCellValue("first_name");
		row.createCell(2).setCellValue("last_name");
		row.createCell(3).setCellValue("birth_day");
		row.createCell(4).setCellValue("sex");
		row.createCell(5).setCellValue("salary");
//		row.createCell(6).setCellValue("super_id");
//		row.createCell(7).setCellValue("branch_id");
		int r = 1;
		while(result.next()) {
			
			int id = result.getInt("emp_id");
			String f_name  = result.getString("first_name");
			String l_name  = result.getString("first_name");
			String b_day  = result.getString("first_name");
			String gender  = result.getString("sex");
			int salary = result.getInt("salary");
			XSSFRow cur_row =  sheet.createRow(r++);
			cur_row.createCell(0).setCellValue(id);
			cur_row.createCell(1).setCellValue(f_name);
			cur_row.createCell(2).setCellValue(l_name);
			cur_row.createCell(3).setCellValue(b_day);
			cur_row.createCell(4).setCellValue(gender);
			cur_row.createCell(5).setCellValue(salary);
			
		}
		FileOutputStream fos = new FileOutputStream(".\\DataFiles\\employee.xlsx");
		workbook.write(fos);
		fos.close();
		con.close();
		System.out.println("File Writen Success");
	}
}
