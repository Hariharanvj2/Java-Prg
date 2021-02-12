import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Date;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Excel {
	public static void main(String[] args) throws IOException {
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("HARI");
		CreationHelper creationHelper = wb.getCreationHelper();
		DataFormat createDataFormat = creationHelper.createDataFormat();
		short format = createDataFormat.getFormat("m/d/yy");
		Row row = sheet.createRow(0);
		CellStyle createCellStyle = wb.createCellStyle();
		createCellStyle.setDataFormat(format);
		row.createCell(0).setCellStyle(createCellStyle);
		row.getCell(0).setCellValue(new java.util.Date());
		row.createCell(1).setCellValue(200);
		Row row2 = sheet.createRow(1);
		row2.createCell(0).setCellValue(100);
		row2.createCell(1).setCellValue(200);
		FileOutputStream in = new FileOutputStream("./Workbook/hari.xlsx");
		wb.write(in);
		FileInputStream read = new FileInputStream("./Workbook/hari.xlsx");
		Sheet sheetAt = wb.getSheetAt(0);
		DataFormatter formatter = new DataFormatter();
		Row row3 = sheetAt.getRow(0);
		for (Row r : sheetAt) {
			for (Cell c : r) {
				String formatCellValue = formatter.formatCellValue(c);
				System.out.print(formatCellValue + "\t");
			}
			System.out.println();
		}

	}
}
