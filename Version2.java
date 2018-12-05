import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Version2 {
	public static void main(String[] args) throws Exception {

		File arch = new File("C:\\Users\\jorge\\Desktop\\SQL.txt");
		BufferedReader br = new BufferedReader(new FileReader(arch));
		String txt;
		String t = "";
		String ts = "";
		String ord = "";
		while ((txt = br.readLine()) != null) {
			ord = ord +" "+ txt;
			if (txt.contains(".") && txt.contains("com") && txt.contains("=")) {
				t = t + " " + txt.substring(txt.indexOf(".") + 1, txt.indexOf("=") - 1) + " "
						+ txt.substring(txt.lastIndexOf(".") + 1, txt.length());
				ts = ts + " " + txt.substring(3, txt.indexOf(".")) + " "
						+ txt.substring(txt.indexOf("=") + 1, txt.lastIndexOf("."));
			} else if (txt.contains(".") && !txt.endsWith(",")) {
				if (txt.contains("WHERE")) {
					t = t + " " + txt.substring(txt.indexOf(".") + 1, txt.indexOf("=") - 1);
					ts = ts + " " + txt.substring(txt.indexOf(" "), txt.indexOf("."));
				} else if (txt.contains("ORDER")) {

				} else {
					ts = ts + " " + txt.substring(0, txt.indexOf("."));
					t = t + " " + txt.substring(txt.indexOf('.') + 1, txt.length());
				}
			} else if (txt.contains(".") && txt.contains(",")) {
				t = t + " " + txt.substring(txt.indexOf('.') + 1, txt.length() - 1);
				ts = ts + " " + txt.substring(0, txt.indexOf("."));

			}
		}
		StringTokenizer st = new StringTokenizer(t);
		StringTokenizer st2 = new StringTokenizer(ts);
		int tm = st2.countTokens();
		String[][] txtar = new String[tm + 1][2];

		for (int j = 1; j < txtar.length; j++) {
			txtar[j][1] = st.nextToken();
			txtar[j][0] = st2.nextToken();
			System.out.println(txtar[j][0] + "   " + txtar[j][1]);
		}
		for (int j = 1; j < txtar.length; j++) {
			if (txtar[j][0].equals("stu")) {
				txtar[j][0] = "Students";
			}
			if (txtar[j][0].equals("com")) {
				txtar[j][0] = "Comments";
			}
		}
		
		System.out.println(ord);
		if (ord.contains("ASC")) {
			txtar[0][0] = "zTableName";
			txtar[0][1] = "zColumnName";
			Arrays.sort(txtar, new Comparator<String[]>() {
				@Override
				public int compare(String[] second, String[] first) {
					// compare the first element
					int comparedTo = first[0].compareTo(second[0]);
					// if the first element is same (result is 0), compare the second element
					if (comparedTo == 0)
						return first[1].compareTo(second[1]);
					else
						return comparedTo;
				}
			});
			txtar[0][0] = txtar[0][0].substring(1);
			txtar[0][1] = txtar[0][1].substring(1);
		}

		if (ord.contains("DST")) {
			txtar[0][0] = "0TableName";
			txtar[0][1] = "0ColumnName";
			Arrays.sort(txtar, new Comparator<String[]>() {
				@Override
				public int compare(String[] first, String[] second) {
					// compare the first element
					int comparedTo = first[0].compareTo(second[0]);
					// if the first element is same (result is 0), compare the second element
					if (comparedTo == 0)
						return first[1].compareTo(second[1]);
					else
						return comparedTo;
				}
			});
			txtar[0][0] = txtar[0][0].substring(1);
			txtar[0][1] = txtar[0][1].substring(1);
		}
		crearExcel("C:\\Users\\jorge\\Desktop\\SQL.xlsx", "SQL", txtar);
	}
	public static void crearExcel(String narch, String nt, String[][] txt) throws FileNotFoundException, IOException {
		Workbook wb = new XSSFWorkbook();
		FileOutputStream fo = new FileOutputStream(narch);
		Sheet h = wb.createSheet(nt);
		
		CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName ( "Arial" );
        font.setBold(true);
        font.setColor((short) 0);
        font.setFontHeightInPoints((short) 25);
        style.setFont(font);
        
		Row[] f = new Row[txt.length];
		Cell[][] c = new Cell[f.length][];
		for (int i = 0; i < f.length; i++) {
			f[i] = h.createRow(i);
			c[i] = new Cell[txt[i].length];
			for (int j = 0; j < c[i].length; j++) {
				c[i][j] = f[i].createCell(j);
				c[i][j].setCellValue(txt[i][j]);
				c[i][j].setCellStyle(style);
			}
		}
		
		h.autoSizeColumn(0);
		h.autoSizeColumn(1);
		
		wb.write(fo);
		
		fo.close();
		System.out.println("Se ha exportado de forma correcta.");
	}

}