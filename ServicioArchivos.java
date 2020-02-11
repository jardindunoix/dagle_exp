package servicios;

import javax.swing.JOptionPane;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import articulos.Articulos;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class ServicioArchivos {
	static String nombreArchivo = "Library_Of_Harvard_.xls";
	static String rutaArchivo = "C:\\Users\\Rodrigo\\Documents\\Projects\\ProjectsEclipse\\Biblioteca\\"
			+ nombreArchivo;
	static String hojaArticulos = "articulos";

	static String hojaUsuarios = "usuarios";

	public static void exportarXLS() {
		ArrayList<Articulos> listaArticulos = new ArrayList<Articulos>();
		ArrayList<Articulos> listaUsuarios = new ArrayList<Articulos>();

		@SuppressWarnings("resource")
		XSSFWorkbook libro = new XSSFWorkbook();
		XSSFSheet hojaArt = libro.createSheet(hojaArticulos);
		XSSFSheet hojaUsu = libro.createSheet(hojaUsuarios);
//		String[] header = new String[] { "Tipo", "Nombre", "Estado", "Código", "CantPág / Duración",
//				"Imprenta / Estudios", "Reservas" };
		CellStyle style = libro.createCellStyle();
		Font font = libro.createFont();
		font.setBold(true);
		font.setColor((short) 54);
		style.setFont(font);

		for (int i = 0; i <= listaArticulos.size(); i++) {
			XSSFRow row = hojaArt.createRow(i);
			for (int j = 0; j < 8; j++) {
				if (i == 0) {
//					XSSFCell cell = row.createCell(j);
//					cell.setCellStyle(style);
//					cell.setCellValue(header[j]);
				} else {
					XSSFCell cell = row.createCell(j);
					if (j == 0) {
						cell.setCellValue(listaArticulos.get(j).getTipo());
					} else if (j == 1) {
						cell.setCellValue(listaArticulos.get(j).getNombre());
					} else if (j == 2) {
						cell.setCellValue(listaArticulos.get(j).getEstado());
					} else if (j == 3) {
						cell.setCellValue(listaArticulos.get(j).getCodigo());
					} else if (j == 4) {
						cell.setCellValue(listaArticulos.get(j).getCantPagDuracion());
					} else if (j == 5) {
						cell.setCellValue(listaArticulos.get(j).getEditorialEstudios());
					} else if (j == 6) {
						cell.setCellValue(listaArticulos.get(j).getReservas());
					}

				}
			}
		}

		File file;
		file = new File(rutaArchivo);
		try (FileOutputStream fileOuS = new FileOutputStream(file)) {
			if (file.exists()) {
				file.delete();
			}

			libro.write(fileOuS);
			fileOuS.flush();
			fileOuS.close();
			JOptionPane.showMessageDialog(null, "Archivo Actualizado!");

		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(null, e);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, e);
		}

	}

}
