package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxCrawler {

	public static void main(String[] args) throws IOException{
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
	
	public static void compare(String filename) throws IOException {
		FileInputStream srcfile = new FileInputStream("./source/" + filename);
		
		XSSFWorkbook wb = new XSSFWorkbook(srcfile);
	}
		

	}


