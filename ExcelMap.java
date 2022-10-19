package Excell.map;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelMap {
	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employee Info");
		XSSFRow row;
		Map<Integer, Object[]> emp = new TreeMap<Integer, Object[]>();
		emp.put(1, new Object[] { "Emp id", "Emp name", "Emp designation" });
		emp.put(2, new Object[] { "011", "Logi", "Developer" });
		emp.put(3, new Object[] { "012", "Raju", "Manager" });
		emp.put(4, new Object[] { "013", "Hari", "CEO" });
		emp.put(5, new Object[] { "014", "JP", "Owner" });

		Set<Integer> keyid = emp.keySet();
		
		int rowid = 0;

		for (Integer key : keyid) {
			row = sheet.createRow(rowid++);
			Object[] objArr = emp.get(key);
			
			int cellid = 0;

			for (Object obj : objArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}

		FileOutputStream out = new FileOutputStream(new File("D:\\Excel Post\\Bank.xlsx"));

		workbook.write(out);
		out.close();
		System.out.println("Excell generated");

	}

}
