package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

/*
 	Here we have to write the data into the excel file using the Poi API
 	
 	So our first move will be to create the data in the form of objects and then create 
 	
 	Workbook->Sheet->Row->Cell
 */

public class WritingExcel {
	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(); // workbook is created

		XSSFSheet sheet = workbook.createSheet("Emp_Info"); // in the given workbook we have created the sheet

		Object empdata[][] = { { "EmpID", "Name", "Job" }, // it is similar to creating a 2D matrix(but here objects)
				{ 101, "Saurav", "WebDev" }, { 102, "Shambhavi", "Fullstack dev" }, { 103, "Sachin", "Teacher" },
				{ 104, "Anjali", "Cisco" }, };

		// Now using the two for loops we can iterate through this 2D object and pushing
		// the data by creating the row and cell in the sheet

		/*
		 * USING FOR LOOP int rows = empdata.length; // no. of rows in the sheet int
		 * cols = empdata[0].length; // no. of colms in sheet
		 * 
		 * for (int r = 0; r < rows; r++) {
		 * 
		 * XSSFRow row = sheet.createRow(r);
		 * 
		 * for (int c = 0; c < cols; c++) {
		 * 
		 * XSSFCell cell = row.createCell(c);
		 * 
		 * Object value = empdata[r][c];
		 * 
		 * if( value instanceof String ) cell.setCellValue((String)(value));
		 * 
		 * if( value instanceof Integer ) cell.setCellValue((Integer)(value));
		 * 
		 * if( value instanceof Boolean ) cell.setCellValue((Boolean)(value)); } }
		 */

		// Now apart from the for loop we can avoid the index stuff bu using the FOREACH
		// LOOP

		int rows = 0;

		for (Object emp[] : empdata) {
			XSSFRow row = sheet.createRow(rows++);
			int cols = 0;
			for (Object value : emp) {
				XSSFCell cell = row.createCell(cols++);

				if (value instanceof String)
					cell.setCellValue((String) (value));

				if (value instanceof Integer)
					cell.setCellValue((Integer) (value));

				if (value instanceof Boolean)
					cell.setCellValue((Boolean) (value));
			}
		}

		// Now we need to write that sheet in our excel file //File HANDLING

		String path = ".\\datafile\\employee.xlsx";

		FileOutputStream outputStream = new FileOutputStream(path);

		workbook.write(outputStream);

		outputStream.close();
		workbook.close();

		System.out.println("Excel file is Written successfully....");

	}

}
