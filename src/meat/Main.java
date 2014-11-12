package meat;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {

    public static void main(String[] args) throws InvalidFormatException,
	    IOException {
	File folder = new File("./target");
	File[] listOfFiles = folder.listFiles();
	FileInputStream inp = null;

	for (File file : listOfFiles) {
	    if (file.getName().toLowerCase().endsWith("xls")
		    || file.getName().toLowerCase().endsWith("xlsx")) {
		inp = new FileInputStream(file);
		System.out.println("File name: " + file.getName());
		Workbook wb = WorkbookFactory.create(inp);
		int sheetCounter = 0;
		Sheet sheet = null;
		String date = null;
		String version = null;
		int added = -1;
		int modified = -1;
		int added2 = -1;
		int modified2 = -1;
		while (true) {
		    try {
			sheet = wb.getSheetAt(sheetCounter++);
		    } catch (Exception e) {
			break;
		    }
		    date = null;
		    version = null;
		    added = -1;
		    modified = -1;
		    added2 = -1;
		    modified = -1;
		    for (int i = 0; i < sheet.getLastRowNum(); ++i) {

			Row currentRow = sheet.getRow(i);
			if (currentRow != null) {
			    Iterator<Cell> cellIterator = currentRow
				    .cellIterator();
			    while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				if (currentCell.getCellType() == Cell.CELL_TYPE_STRING
					&& currentCell.getStringCellValue()
						.toLowerCase().equals("date")) {
				    date = cellIterator.next()
					    .getStringCellValue();
				    break;
				}
				if (currentCell.getCellType() == Cell.CELL_TYPE_STRING
					&& currentCell.getStringCellValue()
						.toLowerCase()
						.equals("version")) {
				    version = cellIterator.next()
					    .getStringCellValue();
				    break;
				}
				if (currentCell.getCellType() == Cell.CELL_TYPE_STRING
					&& currentCell.getStringCellValue()
						.toLowerCase().equals("added")) {
				    Cell next = cellIterator.next();
				    if (added == -1) {
					if (next.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					    added = Double.valueOf(
						    next.getNumericCellValue())
						    .intValue();
					}
					if (next.getCellType() == Cell.CELL_TYPE_STRING) {
					    added = Double.valueOf(
						    next.getStringCellValue())
						    .intValue();
					}
				    } else {
					if (next.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					    added2 = Double.valueOf(
						    next.getNumericCellValue())
						    .intValue();
					}
					if (next.getCellType() == Cell.CELL_TYPE_STRING) {
					    added2 = Double.valueOf(
						    next.getStringCellValue())
						    .intValue();
					}

				    }
				    break;
				}
				if (currentCell.getCellType() == Cell.CELL_TYPE_STRING
					&& currentCell.getStringCellValue()
						.toLowerCase()
						.equals("modified")) {
				    Cell next = cellIterator.next();
				    if (modified == -1) {
					if (next.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					    modified = Double.valueOf(
						    next.getNumericCellValue())
						    .intValue();
					}
					if (next.getCellType() == Cell.CELL_TYPE_STRING) {
					    modified = Double.valueOf(
						    next.getStringCellValue())
						    .intValue();
					}

				    } else {
					if (next.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					    modified2 = Double.valueOf(
						    next.getNumericCellValue())
						    .intValue();
					}
					if (next.getCellType() == Cell.CELL_TYPE_STRING) {
					    modified2 = Double.valueOf(
						    next.getStringCellValue())
						    .intValue();
					}

				    }
				    break;
				}
			    }
			}
		    }
		    if (date != null && version != null && added != -1
			    && modified != -1) {
			System.out.println("Sheet name: "
				+ sheet.getSheetName());
			System.out.println("date: " + date);
			System.out.println("version: " + version);
			System.out.println("added: " + added);
			System.out.println("modified: " + modified);
			System.out.println("added2: " + added2);
			System.out.println("modified2: " + modified2);
			System.out.println("\n");
		    }

		}

	    }
	}
    }

}
