import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	public static void exportDataToExcel(int row, int column, String value) {
		XSSFWorkbook wb = null;
		XSSFSheet sh = null;

		File file = new File("D:\\output.xlsx");

		try {
			FileInputStream fis = new FileInputStream(file);
			wb = new XSSFWorkbook(fis);

			/*
			 * sb=wb.createSheet(); sb=wb.createSheet("nia"); sb=wb.getSheet("nia");
			 */

			if (wb.getSheetIndex("WebTableData") == -1) {
				
				if (row == 0 && column == 0) {
					sh = wb.createSheet("WebTableData");
					System.out.println(sh+" Sheet created");
				} else {
					sh = wb.getSheet("WebTableData");
				}

				if (column == 0) {
					sh.createRow(row).createCell(column).setCellValue(value);
				} else {
					sh.getRow(row).createCell(column).setCellValue(value);
				}

				FileOutputStream fos = new FileOutputStream(file);
				wb.write(fos);

			} else {
				//System.out.println(sh + " Sheet already exists. Hence appending data");
				sh = wb.getSheet("WebTableData");
				if (column == 0) {
					sh.createRow(row).createCell(column).setCellValue(value);
				} else {
					sh.getRow(row).createCell(column).setCellValue(value);
				}

				FileOutputStream fos = new FileOutputStream(file);
				wb.write(fos);
			}
		}

		catch (Exception e) {

			e.printStackTrace();
		}
	}

}
