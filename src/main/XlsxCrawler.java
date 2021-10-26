package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
		}
		

	}


