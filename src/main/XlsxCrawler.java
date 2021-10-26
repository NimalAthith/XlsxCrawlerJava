package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxCrawler {

	public static void main(String[] args) throws Exception{
		// Loading Xlsx
		FileInputStream file = new FileInputStream(new File("./Bom_Compare.xlsx"));
		
		//Workbook
		XSSFWorkbook wb = new XSSFWorkbook(file);
		
		//Sheet
		XSSFSheet ws = wb.getSheetAt(0);
		
		//Iterator
		Iterator<Row> allrow = ws.iterator();
/*		
		while(allrow.hasNext())
		{
			Row row = allrow.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			
			while (cellIterator.hasNext())
			{
				Cell cell = cellIterator.next();
				
				switch (cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue()+ "\t\t\t");
					break;
					
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()+ "\t\t\t");
					break;
					
				default:
					
				
					
				}
				
			}
			System.out.println("");  
			
			
			}
			*/
		
		//List dir
		
				String[] listDir = viewFile();
		
		
	
		//Read column
		
		for (int rowIndex = 1 ; rowIndex <= ws.getLastRowNum(); rowIndex++)
		{
			XSSFRow row = ws.getRow(rowIndex);
			if (row!=null) {
				Cell cell = row.getCell(2);
				if (cell != null) {
					String partname = cell.getStringCellValue();
					System.out.println(partname);
					String filename = matchFile(partname, listDir);
					System.out.println(filename);
					compare(filename);
				}
			}
		}
		
		
		}
	
	public static String[] viewFile() {
		
		File source = new File("./source");
		
		String[] dirList = source.list();
		
		//for(String dir : dirList) {
			//System.out.println(dir);
		//}
		
		return dirList;
		
	}
	
	public static String matchFile(String partname, String[] listDir) {
		
		String filename="";
		
		
		for (String fileInDir : listDir) {
			
			if (fileInDir.startsWith(partname)) {
				//System.out.println(fileInDir);
				filename = fileInDir;
				break;
			}
			
		}
		
		return filename;
	}
	
	public static void compare(String filename) throws Exception {
		FileInputStream srcfile = new FileInputStream("./source/" + filename);
		
		XSSFWorkbook wb = new XSSFWorkbook(srcfile);
		XSSFSheet ws = wb.getSheetAt(0);
		
	//	int rowCount = ws.getPhysicalNumberOfRows();
	/*	
		for (int i = 8; i < rowCount; i++) {
			Row thisRow = ws.getRow(i);
			Iterator<Cell> cellIterator = thisRow.cellIterator();
			Cell cell = cellIterator.next();
			
			if(cell == null || cell.getCellType() == CellType.BLANK) {
				System.out.println("No Difference");
				break;
			}
			
			
			
			
			
		}
		*/
		outerloop:
		for (int rowIndex = 9 ; rowIndex <= ws.getLastRowNum(); rowIndex++) {
			
			XSSFRow row = ws.getRow(rowIndex);
			
			Cell cell = row.getCell(0);
			String out = cell.getStringCellValue();
			boolean va = true;
			try{
			va = out != "";
			}
			finally
			{
			System.out.println("Try");
			boolean po;
			po = false;
			
			if(va)
				po = true;
			
			System.out.println(po);
			if (po){
					
				
				System.out.println("No Difference");
				break outerloop;}
			
			}
				
			
			
			
		}
	}
		

	}


