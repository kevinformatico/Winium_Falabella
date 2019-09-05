package Base;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.net.InetAddress;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;





public class BaseTest {
	public static String[][] matriz;
	public static String directorio;
	public static WiniumDriver driver;
	
	
	public static void crearDirectorios() throws Exception {
		
		directorio = crearDirectorio();
	}
	public static void abrirSwitch() throws MalformedURLException, Exception {
		
		
		
		DesktopOptions option = new DesktopOptions();
		option.setApplicationPath("C:\\Users\\Hector Castillo\\AppData\\Roaming\\IBM\\Personal Communications\\switch.WS");
		driver = new WiniumDriver(new URL("http://localhost:9999"), option);

		matriz =  getDataTxt("DataSwitch");
	}
	
	public static void cerrarApp() throws Exception {
		
//		Robot teclado = new Robot();
//		
//		teclado.keyPress(KeyEvent.VK_ALT);
//		teclado.keyPress(KeyEvent.VK_F4);
//		teclado.keyRelease(KeyEvent.VK_F4);
//		teclado.keyRelease(KeyEvent.VK_ALT);
		
//		Process p = Runtime.getRuntime().exec("PCSWS.exe");
//		p.destroy();
		
//		driver.close();
		driver.quit();
	}
	
	public static String leerMatrizTxt(String tag, String[][] matriz) {
		String value = null;
		for (int i = 0; i < matriz.length; i++) {
			if (matriz[i][0].toUpperCase().equals(tag.toUpperCase())) {
				value = matriz[i][1];
				break;
			}
		}
		return value;
	}

	public static String[][] getDataTxt(String caso) throws Exception {
		String computerName = InetAddress.getLocalHost().getHostName();
		String path = "";
		String[][] matriz = null;
		String line = null;
		String[] arrLine;
		int col = 2;
		try {
		if (computerName.contains("f8cloud")) {
			
			path = "F:\\apps\\MaquinasCMR\\desarrollos\\Ejecucion_Automatizada\\Datos_Temporales\\" + caso + ".txt";
		} else {
			path = "C:\\desarrollos\\Ejecucion_Automatizada\\Datos_Temporales\\" + caso + ".txt";
		}		
		
			int filas = cuentaFilasFile(path);
			BufferedReader in = new BufferedReader(new FileReader(path));
			matriz = new String[filas][col];

			for (int i = 0; i < filas; i++) {
				line = in.readLine();
//				System.out.println(line);
				arrLine = line.split("=");
				matriz[i][0] = arrLine[0];
				if (arrLine.length > 1) {
					matriz[i][1] = arrLine[1];
				}

			}
			in.close();
		} catch (Exception e) {
			System.out.println("Error al leer datos de prueba: " + e.getMessage());
		}
		return matriz;
	}

	public static int cuentaFilasFile(String path) {
		int i = 0;
		try {
			BufferedReader in = new BufferedReader(new FileReader(path));
			while ((in.readLine()) != null) {
				i++;
			}
			in.close();
		} catch (Exception e) {
			System.out.println("Error al leer el archivo");
		}
		return i;
	}
	
	
	public static String[][] retornaDatosExcel() throws IOException {
		String[][] matriz = null;
		String rutaProyecto = System.getProperty("user.dir");


			String ruta_xlsx = rutaProyecto +  "\\resources\\DataTrx.xlsx";
			
			try {

				FileInputStream file = new FileInputStream(new File(ruta_xlsx));
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				XSSFSheet sheet = workbook.getSheetAt(0);

				Iterator<Row> rowIterator = sheet.iterator();
				Row row;

				String valorString = "";
				int valorInt = -1;
				int contadorFilas = 0, contadorCol = 0;

				matriz = new String[LeerExcel.countRow(sheet) + 1][LeerExcel.countCol(sheet)];

				while (rowIterator.hasNext()) {

					row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					Cell celda;

					while (cellIterator.hasNext()) {

						celda = cellIterator.next();

						switch (celda.getCellType()) {

						case Cell.CELL_TYPE_NUMERIC:
							valorInt = (int) Math.round(celda.getNumericCellValue());
							break;

						case Cell.CELL_TYPE_STRING:
							valorString = celda.getStringCellValue();
							break;
						}

						if (!valorString.equals("")) {
							matriz[contadorFilas][contadorCol] = valorString;
							valorString = "";
						} else if (valorInt != -1) {
							matriz[contadorFilas][contadorCol] = Integer.toString(valorInt);
							valorInt = -1;
						}

						if (contadorCol != LeerExcel.countCol(sheet) - 1) {
							contadorCol++;
						} else {
							contadorCol = 0;
						}
					}
					contadorFilas++;

				}

				if (matriz.length < 2) {
					System.out.println("El archivo excel no contiene datos");
					// workbook.close();
				}

				// workbook.close();
			} catch (Exception e) {
				System.out.println("El archivo excel no existe");
				// System.exit(0);
			}

			return matriz;
}

	
	public static String crearDirectorio() throws IOException {
		String str = "";
		try {

				String anoMes, dia, horaEjecucucion;
				String rutaProyecto = System.getProperty("user.dir");
				String path_screenshot = rutaProyecto + "\\Evidencia\\";
				Date objDate = new Date();
				anoMes = "MMyyyy";
				dia = "dd";
				horaEjecucucion = "hhmmss";

				SimpleDateFormat anoFormat = new SimpleDateFormat(anoMes);
				String AnoMesFormateado = anoFormat.format(objDate);
				SimpleDateFormat diaFormat = new SimpleDateFormat(dia);
				String diaFormateado = diaFormat.format(objDate);
				SimpleDateFormat horaFormat = new SimpleDateFormat(horaEjecucucion);
				String HoraFormateada = horaFormat.format(objDate);

				File directorioAnoMes = new File(path_screenshot + AnoMesFormateado);
				File directorioDia = new File(directorioAnoMes + "/" + diaFormateado);
				File directorioHora = new File(directorioDia + "/" + HoraFormateada);

				if (!directorioAnoMes.exists()) {
					directorioAnoMes.mkdir();
				}
				if (!directorioDia.exists()) {
					directorioDia.mkdir();
				}
				if (!directorioHora.exists()) {
					directorioHora.mkdir();
				}

				str = path_screenshot + AnoMesFormateado + "/" + diaFormateado + "/" + HoraFormateada;

			

		} catch (Exception e) {
			System.out.println(e.getMessage().toString());

		}
		return str;
	}

	
}
