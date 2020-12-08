package MavenTest.ExcelData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class AccessData {

	public static void main(String[] args) throws IOException {
		
		/* ArrayList<String> tcdata = new ArrayList<String>();
		
		// Read testcase data from an Excel file so create an input stream
		FileInputStream fis = new FileInputStream("C://Users//Z_Wong 25//Documents//udemy//selenium//dataset.xlsx");
	
		// create an Excel workbook object and give it the input stream
		XSSFWorkbook workbk = new XSSFWorkbook(fis);
		
		// now find out the number of sheets in the workbook
		int numsheets = workbk.getNumberOfSheets();
		
		// find the sheet with the testcases and it is called CompanySet
		for (int i = 0; i < numsheets; i++) {
			if (workbk.getSheetName(i).equalsIgnoreCase("CompanySet")) {
				// we have the correct sheet so now look at the header row
				XSSFSheet sheet = workbk.getSheetAt(i);
				
				// so from this sheet access the first row which is the header
				Iterator<Row> arow = sheet.iterator();
				
				if (arow.hasNext()) {
					Row firstrow = arow.next();
								
					// now look at each cell to find the testcase column
					Iterator<Cell> cel = firstrow.iterator();
					
					int index = 0;
					int column = 0;
					
					while (cel.hasNext()) {
						// find out what the header cell is
						Cell celname = cel.next();
						
						if (celname.getStringCellValue().equalsIgnoreCase("TestCase")) {
							
							// okay now look for a certain testcase
							column = index;
						}
						index++;
					}
					// okay now let's find a particular testcase and use its data
					while (arow.hasNext()) {
						Row nextrow = arow.next();				
						
						if (nextrow.getCell(column).getStringCellValue().equalsIgnoreCase("TC4")) {
							// we found the correct testcase so now get the data
							
							Iterator<Cell> dat = nextrow.cellIterator();
							
							while (dat.hasNext()) {
								Cell thedat = dat.next();
							
								tcdata.add(thedat.getStringCellValue());
							}
							
							// return the array list to calling class
							//return(tcdata);
						}
					}
				}
				
				
				
			}
		}
		return(tcdata); */
	}
	
	public ArrayList<String> getData(String tcreq) throws IOException {
	ArrayList<String> tcdata = new ArrayList<String>();
	
	// Read testcase data from an Excel file so create an input stream
	FileInputStream fis = new FileInputStream("C://Users//Z_Wong 25//Documents//udemy//selenium//dataset.xlsx");

	// create an Excel workbook object and give it the input stream
	XSSFWorkbook workbk = new XSSFWorkbook(fis);
	
	// now find out the number of sheets in the workbook
	int numsheets = workbk.getNumberOfSheets();
	
	// find the sheet with the testcases and it is called CompanySet
	for (int i = 0; i < numsheets; i++) {
		if (workbk.getSheetName(i).equalsIgnoreCase("CompanySet")) {
			// we have the correct sheet so now look at the header row
			XSSFSheet sheet = workbk.getSheetAt(i);
			
			// so from this sheet access the first row which is the header
			Iterator<Row> arow = sheet.iterator();
			
			if (arow.hasNext()) {
				Row firstrow = arow.next();
							
				// now look at each cell to find the testcase column
				Iterator<Cell> cel = firstrow.iterator();
				
				int index = 0;
				int column = 0;
				
				while (cel.hasNext()) {
					// find out what the header cell is
					Cell celname = cel.next();
					
					if (celname.getStringCellValue().equalsIgnoreCase("TestCase")) {
						
						// okay now look for a certain testcase
						column = index;
					}
					index++;
				}
				// okay now let's find a particular testcase and use its data
				while (arow.hasNext()) {
					Row nextrow = arow.next();				
					
					if (nextrow.getCell(column).getStringCellValue().equalsIgnoreCase(tcreq)) {
						// we found the correct testcase so now get the data
						
						Iterator<Cell> dat = nextrow.cellIterator();
						
						while (dat.hasNext()) {
							Cell thedat = dat.next();
							
													
							org.apache.poi.ss.usermodel.CellType typ = thedat.getCellType();
						
							if (thedat.getCellType() == typ.STRING) {
								tcdata.add(thedat.getStringCellValue());
							} else
							//	if (thedat.getCellType() == typ.NUMERIC) {
								// turn it into a string
								tcdata.add(NumberToTextConverter.toText(thedat.getNumericCellValue()));
							//		tcdata.add((int) thedat.getNumericCellValue());
							//}
						}
						
						// return the array list to calling class
						//return(tcdata);
					}
				}
			}
			
			
			
		}
	}
	
	workbk.close();
	return(tcdata);	
	}

}
