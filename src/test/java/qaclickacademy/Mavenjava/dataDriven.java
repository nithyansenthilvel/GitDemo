package qaclickacademy.Mavenjava;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	// Identify the TestCase column by scanning the entire 1st row
	// Once column is identified then scan entire TestCase column to identify
	// purchase TestCase row

	public ArrayList<String> getData(String testCaseName) throws IOException {
		// File input stream argument
		ArrayList<String> a = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C:\\Users\\nithy\\Desktop\\Test.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i); // Identify TestCase column by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator(); // Sheet is collection of rows
				Row firstRow = rows.next();
				Iterator<Cell> ce = firstRow.cellIterator(); // Row is collection of cells
				int k = 0;
				int column = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column = k;

					}
					k++;
				}
				System.out.println(column);
				// Once column is identified then scan entire testcase column to identify
				// purchase testCase row
				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						// after you grab purchase testCase row = pull all the data of that row and feed
						// into test

						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							Cell c = cv.next();
							if (c.getCellTypeEnum() == CellType.STRING) {
								a.add(c.getStringCellValue());
							} else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}

					}
				}
			}
		}
		return a;
	}
}
