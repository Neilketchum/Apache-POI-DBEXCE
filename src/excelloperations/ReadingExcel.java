package excelloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {
	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception{
		String excelFilePath = ".\\DataFiles\\country.xlsx";
		FileInputStream fos = new FileInputStream(excelFilePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fos);
		 XSSFSheet sheet  = workbook.getSheetAt(0);
//		 Using For Loop
//		 int rows = sheet.getLastRowNum();
//		 int cols = sheet.getRow(0).getLastCellNum();
//		 for(int r = 0;r<rows;r++) {
//			 XSSFRow currentRow = sheet.getRow(r);
//			 for(int c = 0;c<cols;c++) {
//				XSSFCell currentCell =  currentRow.getCell(c);
//			 	switch (currentCell.getCellType()) {
//				case STRING: {
//					System.out.print(currentCell.getStringCellValue() + "----------");
//					break;
//				}
//				case NUMERIC:{
//					System.out.print(currentCell.getNumericCellValue() + "----------");
//					break;
//				}
//				
//				default:
//					throw new IllegalArgumentException("Unexpected value: " + currentCell.getCellType());
//				}
//			 }
//			 System.out.println();
//		 }
//		Using Iterator
		 Iterator iter = sheet.iterator();
		 while(iter.hasNext()) {
			 XSSFRow row = (XSSFRow)iter.next();
		
			 Iterator celIterator = row.cellIterator();
			 while(celIterator.hasNext()) {
				 XSSFCell cell = (XSSFCell)celIterator.next();
				 switch (cell.getCellType()) {
					case STRING: {
						System.out.print(cell.getStringCellValue() + "----------");
						break;
					}
					case NUMERIC:{
						System.out.print(cell.getNumericCellValue() + "----------");
						break;
					}
					
					default:
						throw new IllegalArgumentException("Unexpected value: " + cell.getCellType());
					}
				 }
				 System.out.println();
			 }
			 
		 }
		 
		
	}
