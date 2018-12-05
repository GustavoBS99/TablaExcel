import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TablaExcel {
	public static void main(String[] args) throws Exception {

		File arch = new File("C:\\Users\\jorge\\Desktop\\SQL.txt");
		BufferedReader br = new BufferedReader(new FileReader(arch));

		String linea;
		int lincant = 0;
		String texto = "";
		String texto2 = "";
		String nomtab1 = "Students";
		String nomtab2 = "Comments";

		while ((linea = br.readLine()) != null) {
			lincant++;
			if (linea.contains(".") && linea.contains("com") && linea.contains("=")) {
				texto = texto + " " + linea.substring(linea.indexOf(".") + 1, linea.indexOf("=") - 1) + " "
						+ linea.substring(linea.lastIndexOf(".") + 1, linea.length());
				texto2 = texto2 + " " + linea.substring(3, linea.indexOf(".")) + " "
						+ linea.substring(linea.indexOf("=") + 1, linea.lastIndexOf("."));

			} else if (linea.contains(".") && !linea.endsWith(",")) {
				if (linea.contains("WHERE")) {
					texto = texto + " " + linea.substring(linea.indexOf(".") + 1, linea.indexOf("=") - 1);
					texto2 = texto2 + " " + linea.substring(linea.indexOf(" "), linea.indexOf("."));
				} else if (linea.contains("ORDER")) {
					texto = texto + " " + linea.substring(linea.indexOf(".") + 1, linea.lastIndexOf(" "));
					texto2 = texto2 + " " + linea.substring(linea.indexOf("Y ") + 1, linea.indexOf("."));

				} else {
					texto2 = texto2 + " " + linea.substring(0, linea.indexOf("."));
					texto = texto + " " + linea.substring(linea.indexOf('.') + 1, linea.length());
				}
			} else if (linea.contains(".") && linea.contains(",")) {
				texto = texto + " " + linea.substring(linea.indexOf('.') + 1, linea.length() - 1);
				texto2 = texto2 + " " + linea.substring(0, linea.indexOf("."));

			}else if (linea.equals("FROM") || linea.equals("SELECT")) {

			} else if (linea.contains("JOIN")) {

			} else {

				texto = texto + " ";
				
				texto2 = texto2 + " " ;
			}

			
			
		

		}
		System.out.println(texto2);
		System.out.println(lincant);
		
		System.out.println(texto);
		StringTokenizer st = new StringTokenizer(texto);
		StringTokenizer st2 = new StringTokenizer(texto2);
		int tok = st.countTokens();

		String aux = null;
		int tok2 = st2.countTokens();

		System.out.println(texto2);
		String[][] textoarray = new String[tok2 + 1][2];

		textoarray[0][0] = "TableName";
		textoarray[0][1] = "ColumnName";
		System.out.println(tok);
		System.out.println(tok2);
		int c = 0;
		for (int j = 1; j < tok; j++) {

			textoarray[j][1] = st.nextToken();
			textoarray[j][0] = st2.nextToken();
			System.out.println(textoarray[j][0] + "   " + textoarray[j][1]);
			

		}
		System.out.println(nomtab2);


		for (int j = 1; j < tok; j++) {
			if (textoarray[j][0].equals("stu")) {
				textoarray[j][0] = "Students";

			} 
			if (textoarray[j][0].equals("com")) {
				textoarray[j][0] = "Comments";

			}

		}

		System.out.println(nomtab1+"nomba");
		
		System.out.println(nomtab2+"asdasd");

		crearExcel("C:\\Users\\tavis\\Desktop\\SQL.xlsx", "ex", textoarray);
	}

	public static void crearExcel(String fileName, String tabName, String[][] data)
			throws FileNotFoundException, IOException

	{
		// Create new workbook and tab
		Workbook wb = new XSSFWorkbook();
		FileOutputStream fileOut = new FileOutputStream(fileName);
		Sheet sheet = wb.createSheet(tabName);

		// Create 2D Cell Array
		Row[] fila = new Row[data.length];
		Cell[][] celda = new Cell[fila.length][];

		// Define and Assign Cell Data from Given
		for (int i = 0; i < fila.length; i++) {
			fila[i] = sheet.createRow(i);
			celda[i] = new Cell[data[i].length];

			for (int j = 0; j < celda[i].length; j++) {
				celda[i][j] = fila[i].createCell(j);
				celda[i][j].setCellValue(data[i][j]);
			}

		}

		// Export Data
		wb.write(fileOut);
		fileOut.close();
		System.out.println("File exported successfully");
	}

}