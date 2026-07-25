package com.incesoft.tools.excel.xlsx;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import javax.xml.stream.XMLStreamException;

import com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader;
import com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook.Commiter;

/**
 * @author floyd
 * 
 */
public class TestSJXLSX {

	public static void addStyleAndRichText(SimpleXLSXWorkbook wb, Sheet sheet)
			throws Exception {
		Font font2 = wb.createFont();
		font2.setColor("FFFF0000");
		Fill fill = wb.createFill();
		fill.setFgColor("FF00FF00");
		CellStyle style = wb.createStyle(font2, fill);

		RichText richText = wb.createRichText();
		richText.setText("test_text");
		Font font = wb.createFont();
		font.setColor("FFFF0000");
		richText.applyFont(font, 1, 2);
		sheet.modify(2, 27, (String) null, style);
		sheet.modify(2, 27, richText, null);
	}

	static public void addRecordsOnTheFly(SimpleXLSXWorkbook wb, Sheet sheet,
			int rowOffset) {
		int columnCount = 52;
		int rowCount = 10;
		int offset = rowOffset == -1 ? sheet.getModfiedRowLength() : rowOffset;
		for (int r = offset; r < offset + rowCount; r++) {
			for (int c = 0; c < columnCount; c++) {
				sheet.modify(r, c, r + "," + c, null);
			}
		}
	}

	private static void printRow(int rowPos, Cell[] row) {
		int cellPos = -1;
		for (Cell cell : row) {
			cellPos++;
			if (cell == null)
				continue;
			System.out.println(Sheet.getCellId(rowPos, cellPos) + "="
					+ cell.getValue());
		}
	}

	public static void testLoadALL(SimpleXLSXWorkbook workbook) {
		// medium data set,just load all at a time
		Sheet sheetToRead = workbook.getSheet(0);
		List<Cell[]> rows = sheetToRead.getRows();
		int rowPos = 0;
		for (Cell[] row : rows) {
			printRow(rowPos, row);
			rowPos++;
		}
	}

	public static void testIterateALL(SimpleXLSXWorkbook workbook) {
		// here we assume that the sheet contains too many rows which will leads
		// to memory overflow;
		// So we get sheet without loading all records
		Sheet sheetToRead = workbook.getSheet(0, false);
		SheetRowReader reader = sheetToRead.newReader();
		Cell[] row;
		int rowPos = 0;
		while ((row = reader.readRow()) != null) {
			printRow(rowPos, row);
			rowPos++;
		}
	}

	public static void testWrite(SimpleXLSXWorkbook workbook,
			OutputStream outputStream) throws Exception {
		Sheet sheet = workbook.getSheet(0);
		addRecordsOnTheFly(workbook, sheet, 0);
		workbook.commit(outputStream);
	}

	/**
	 * Commit serveral times for large data set
	 * 
	 * @param workbook
	 * @param output
	 * @throws Exception
	 */
	public static void testWriteByIncrement(SimpleXLSXWorkbook workbook,
			OutputStream output) throws Exception {
		Commiter commiter = workbook.newCommiter(output);
		commiter.beginCommit();

		Sheet sheet = workbook.getSheet(0, false);
		commiter.beginCommitSheet(sheet);
		addRecordsOnTheFly(workbook, sheet, sheet.getModfiedRowLength());
		commiter.commitSheetWrites();
		// ...serveral empty rows...
		addRecordsOnTheFly(workbook, sheet, 20);
		commiter.commitSheetWrites();
		addRecordsOnTheFly(workbook, sheet, sheet.getModfiedRowLength());
		commiter.commitSheetWrites();
		commiter.endCommitSheet();

		commiter.endCommit();
	}

	/**
	 * first, modify the original sheet; and then append some data
	 * 
	 * @param workbook
	 * @param output
	 * @throws Exception
	 */
	public static void testMergeBeforeWrite(SimpleXLSXWorkbook workbook,
			OutputStream output) throws Exception {
		Sheet sheet = workbook.getSheet(0, true);// assuming original data
		// set is large
		// comments
		addStyleAndRichText(workbook, sheet);
		addRecordsOnTheFly(workbook, sheet, 5);

		sheet.modify(2, 27, "0..0 comment");
		sheet
				.modify(1, 1,
						"0..0 comment\n0..0 comment\n0..0 comment\n0..0 comment\n0..0 comment\n");

		Commiter commiter = workbook.newCommiter(output);
		commiter.beginCommit();
		commiter.beginCommitSheet(sheet);
		// merge it first,otherwise the modification will not take effect
		commiter.commitSheetModifications();

		// start to write
		// row = -1, for appending after the last row
		int rowLenModified = sheet.getModfiedRowLength();
		sheet.modify(rowLenModified, 1, "append1", null);
		sheet.modify(rowLenModified, 2, "append2", null);
		// lets assume there are many rows here...
		commiter.commitSheetWrites();// flush writes,save memory

		rowLenModified = sheet.getModfiedRowLength();
		sheet.modify(rowLenModified, 1, "append3", null);
		sheet.modify(rowLenModified, 2, "append4", null);
		// lets assume there are many rows here,too ...
		commiter.commitSheetWrites();// flush writes,save memory

		commiter.endCommitSheet();
		commiter.endCommit();
	}

	private static SimpleXLSXWorkbook newWorkbook() {
		return new SimpleXLSXWorkbook(new File("src/main/resources/new_retail_data.xlsx"));
//		return new SimpleXLSXWorkbook(new File("src/main/resources/1.xlsx"));
		// return new SimpleXLSXWorkbook(new File("/sample_no_rel.xlsx"));
	}

	private static OutputStream newOutput(String filename) throws FileNotFoundException {
//		return new BufferedOutputStream(new FileOutputStream(new File("src/main/resources/" + filename + ".xlsx")));
		return new BufferedOutputStream(new FileOutputStream(new File(filename + ".xlsx")));
	}

	public static void sortSheetByColumn(SimpleXLSXWorkbook workbook, int sheetIndex, String sheetname, final int... sortColumns) {
		Sheet sheet = workbook.getSheet(sheetIndex);
		sheet.sort(sortColumns);
		
		try {
			OutputStream output = newOutput("src/main/resources/sorted_output");
			Commiter commiter = workbook.newCommiter(output);
			commiter.beginCommit();
			if(!sheetname.isEmpty()) {
				sheet.setSheetName(sheetname);
			}
			commiter.beginCommitSheet(sheet);
			// merge it first,otherwise the modification will not take effect
			commiter.commitSheetModifications();
			commiter.commitSheetWrites();// flush writes,save memory

			commiter.endCommitSheet();
			commiter.endCommit();
			output.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (XMLStreamException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	

	/**
	 * Test method to demonstrate setting sheet names
	 */
	public static void testSetSheetName(SimpleXLSXWorkbook workbook, OutputStream output) throws Exception {
		Sheet sheet = workbook.getSheet(0);
		// Set the sheet name
		sheet.setSheetName("MyDataSheet");
		addRecordsOnTheFly(workbook, sheet, 0);
		workbook.commit(output);
	}

	public static void main(String[] args) throws Exception {
		// READ by classic mode - load all records
		long st = System.currentTimeMillis();
//		for (int i = 0; i < 1; i++) {
			SimpleXLSXWorkbook workbook = newWorkbook();
			List<Sheet> sheets = workbook.sheets;
			for(Sheet s : sheets) {
//				System.out.println("Shhet name ::::"+s.getSheetName());
			}
			sortSheetByColumn(workbook, 0,  "test output", new int[] { 0,2 });
			
			// Save the workbook with sorted data to output file
//			OutputStream output = newOutput("sorted_output");
//			workbook.commit(output);
//			output.close();
//			
//			testLoadALL(workbook);
//			inMemoryIterateALL(workbook);
			workbook.close();
//		}

		// READ by stream mode - iterate records one by one
		// testIterateALL(newWorkbook());
		System.out.println("Excel process time taken ::: "+(System.currentTimeMillis() - st)/1000 + " seconds");
		// //
		// // // WRITE - we take WRITE as a special kind of MODIFY
//		OutputStream output = newOutput("write");
//		testWrite(newWorkbook(), output);
//		output.close();
		// //
		// // // WRITE large data
		// output = newOutput("write_inc");
		// testWriteByIncrement(newWorkbook(), output);
		// output.close();
		//
		// // MODIFY it and WRITE large data
//		output = newOutput("merge_write");
//		testMergeBeforeWrite(newWorkbook(), output);
//		output.close();
	}
}
