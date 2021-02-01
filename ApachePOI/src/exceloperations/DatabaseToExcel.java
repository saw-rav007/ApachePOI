package exceloperations;

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

public class DatabaseToExcel {

	public static void main(String[] args) throws SQLException, IOException {

		// connection
		Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/mydb", "root", "sau123rav");

		// Statement query

		Statement st = con.createStatement();
		ResultSet rs = st.executeQuery("select * from student");

		// Excel

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("STUDENTS INFO");

		XSSFRow row = sheet.createRow(0);

		// Setting the header of the excel

		row.createCell(0).setCellValue("ROLL NO.");
		row.createCell(1).setCellValue("NAME");

		int r = 1;
		while (rs.next()) {
			int roll = rs.getInt("ROLL");
			String name = rs.getString("NAME");

			row = sheet.createRow(r++);

			row.createCell(0).setCellValue(roll);
			row.createCell(1).setCellValue(name);
		}

		// FILE WRITE OPERATIONS

		String path = ".\\datafile\\student.xlsx";
		FileOutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		System.out.println("DB data written successfully in the excel file");

		outputStream.close();
		workbook.close();

	}
}
