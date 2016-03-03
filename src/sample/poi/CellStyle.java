package sample.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellStyle {
	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("cellstyle");
		XSSFRow row = spreadsheet.createRow(1);
		row.setHeight((short)800);
		XSSFCell cell = row.createCell(1);
		cell.setCellValue("test of merging");
		
		// MEARGING CELLS
		spreadsheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 4));
		
		// CELL Alignment
		row = spreadsheet.createRow(5);
		cell = row.createCell(0);
		row.setHeight((short)800);
		// Top Left alignment
		XSSFCellStyle style1 = workbook.createCellStyle();
		spreadsheet.setColumnWidth(0, 8000);
		style1.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		style1.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
		cell.setCellValue("Top Left");
		cell.setCellStyle(style1);
		
		row = spreadsheet.createRow(6);
		cell = row.createCell(1);
		row.setHeight((short)800);
		// Center Align Cell Contents
		XSSFCellStyle style2 = workbook.createCellStyle();
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cell.setCellValue("Center Aligned");
		cell.setCellStyle(style2);
		
		row = spreadsheet.createRow(7);
		cell = row.createCell(2);
		row.setHeight((short)800);
		// Bottom Right alignment
		XSSFCellStyle style3 = workbook.createCellStyle();
		style3.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_BOTTOM);
		cell.setCellValue("Bottom Right");
		cell.setCellStyle(style3);
		
		row = spreadsheet.createRow(8);
		
		FileOutputStream out = new FileOutputStream(new File("cellstyle.xlsx"));
		workbook.write(out);
		workbook.close();
		System.out.println("cellstyle.xlsx written successfully");
	}
}
