package sample.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Formula {

	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("formula");
		
		FileOutputStream out = new FileOutputStream(new File("formula.xlsx"));
		workbook.write(out);
		out.close();
		workbook.close();
		System.out.println("formula.xlsx written successfully");
	}

}
