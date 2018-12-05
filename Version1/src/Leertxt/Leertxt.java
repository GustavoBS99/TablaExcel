package Leertxt;

import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class Leertxt {
	public static void main(String[] args) {

		LinkedList<String[]> text_lines = new LinkedList<>();
		try (BufferedReader br = new BufferedReader(new FileReader("C:\\Users\\tavis\\Desktop\\SQL.txt"))) {
			String sCurrentLine;
			while ((sCurrentLine = br.readLine()) != null) {
				text_lines.add(sCurrentLine.split("\\."));
	
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

		String fileName = "C:\\Users\\tavis\\Desktop\\SQL.xls";
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Test");
		int row_num = 0;
		for (String[] line : text_lines) {
			Row row = sheet.createRow(row_num++);
			int cell_num = 0;
			for (String value : line) {
				Cell cell = row.createCell(cell_num++);
				cell.setCellValue(value);
			}
		}

		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(fileName);
			workbook.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	
	}
}