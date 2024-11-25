package Programs;  

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadAndWrite {
	public static void main(String[] args) throws IOException {

		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet("Sheet1");
		Object[][] date = { { "Name", "Age", "Email" }, { "John Doe", "30", "john@test.com" },
				{ "John Doe", "28", "john@test.com" }, { "Bob Smith", "35", "jacky@example.com" },
				{ "Swapnil", "37", "swapnil@example.com" }

		};
		int rowNum = 0;
		for (Object[] rowData : date) {
			Row row = sheet.createRow(rowNum++);
			int columnNum = 0;
			for(Object field : rowData) {
				Cell cell = row.createCell(columnNum++);
				if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
			}

		}
		
		FileOutputStream fileOut = new FileOutputStream("ExcelReadAndWrite.xlsx");
        book.write(fileOut);
        fileOut.close();
		
		FileInputStream fileIn = new FileInputStream("ExcelReadAndWrite.xlsx");
		Workbook readWorkbook = new XSSFWorkbook(fileIn);
        Sheet readSheet = readWorkbook.getSheetAt(0);
        for (Row row : readSheet) {
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t");
                        break;
                    default:
                        System.out.print("Unknown\t");
                }
            }
            System.out.println(); // Move to the next line after each row
        }

        readWorkbook.close();
        fileIn.close();
		
	}
}