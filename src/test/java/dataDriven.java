import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
	
	ArrayList<String> a = new ArrayList<String>();

	public ArrayList getData(String testCaseName) throws IOException {

		FileInputStream fls = new FileInputStream("/Users/user/Desktop/Selenium/demoData.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fls);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testData")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				java.util.Iterator<Row> rows = sheet.iterator();

				Row firstrow = rows.next();

				java.util.Iterator<Cell> ce = firstrow.cellIterator();

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

				while (rows.hasNext()) {
					Row r = rows.next();

					if (r.getCell(column).getStringCellValue().equalsIgnoreCase("Add Profile")) {
						// after you grab Purchase testCase row = pull all the data row and feed into
						// test

						java.util.Iterator<Cell> cv = r.cellIterator();

						while (cv.hasNext()) {
							
							Cell c = cv.next();
							
							if (c.getCellType() == CellType.STRING)
							{
							a.add(c.getStringCellValue());
							}
							else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}

				}
				
				
			}
		}
		return a;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	}

}
