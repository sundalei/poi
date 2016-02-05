package sample.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TypesofCells {
	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("cell types");

		XSSFRow row = sheet.createRow(2);
		row.createCell(0).setCellValue("Type of Cell");
		row.createCell(1).setCellValue("cell value");

		row = sheet.createRow(3);
		row.createCell(0).setCellValue("set cell type BLANK");
		row.createCell(1);
		
		row = sheet.createRow(4);
		row.createCell(0).setCellValue("set cell type BOOLEAN");
		row.createCell(1).setCellValue(true);
		
		row = sheet.createRow(5);
		row.createCell(0).setCellValue("set cell type ERROR");
		row.createCell(1).setCellValue(XSSFCell.CELL_TYPE_ERROR);
		
		row = sheet.createRow(6);
		row.createCell(0).setCellValue("set cell type date");
		row.createCell(1).setCellValue(new Date());
		
		row = sheet.createRow(7);
		row.createCell(0).setCellValue("set cell type numeric");
		row.createCell(1).setCellValue(20);
		
		row = sheet.createRow(8);
		row.createCell(0).setCellValue("set cell type string");
		row.createCell(1).setCellValue("A String");

		FileOutputStream out = new FileOutputStream(new File("typesofcells.xlsx"));
		workbook.write(out);
		out.close();
		workbook.close();
		System.out.println("typesofcells.xlsx written successfully");
	}
}
