package Page;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;

import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

import Base.BaseTest;

public class winium_page {

	public static void abrirAplicacion(String directorio) throws Exception {

		BaseTest.driver.findElement(By.name("Cancelar")).click();
		BaseTest.driver.findElement(By.name("Maximizar")).click();
		BaseTest.driver.findElement(By.name("Archivo")).click();
		screenshot(directorio);
		BaseTest.driver.findElement(By.name("Abrir...")).click();
		BaseTest.driver.findElement(By.name("No")).click();
		screenshot(directorio);
		BaseTest.driver.findElement(By.name("Abrir")).click();

	}

	public static void loginSwitch(String usuario, String clave, String directorio) throws Exception {

		Thread.sleep(5000);
		Robot teclado = new Robot();
		DesgloceCaracteres(usuario);

		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);

		DesgloceCaracteres(clave);

//		screenshot(directorio);

		for (int i = 0; i < 4; i++) {
			teclado.keyPress(KeyEvent.VK_ENTER);
			teclado.keyRelease(KeyEvent.VK_ENTER);
		}

		DesgloceCaracteres("az7");
		screenshot(directorio);

		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);


	}
	
	public static void consultaTrx(String directorio) throws Exception {
		
		Robot teclado = new Robot();
		DesgloceCaracteres("05");
//		screenshot(directorio);

		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);

		DesgloceCaracteres("04");
//		screenshot(directorio);

		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
	}
	
	public static void consultaCanal(String directorio) throws Exception {
		
		Robot teclado = new Robot();
		DesgloceCaracteres("02");

//		screenshot(directorio);
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
//		screenshot(directorio);
	
	}
	
	public static void consultaSatif(String directorio) throws Exception {
		
		Robot teclado = new Robot();
		DesgloceCaracteres("02");

		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		for (int i = 0; i < 2; i++) {
			teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
			teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		}
		
		for (int i = 0; i < 11; i++) {
			teclado.keyPress(KeyEvent.VK_TAB);
			teclado.keyRelease(KeyEvent.VK_TAB);
		}
		
		teclado.keyPress(KeyEvent.VK_O);
		teclado.keyRelease(KeyEvent.VK_O);
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		DesgloceCaracteres("24");
		
		teclado.keyPress(KeyEvent.VK_F11);
		teclado.keyRelease(KeyEvent.VK_F11);
	
	}
	
	

	public static void navegacionSwitch(String pan, String fechaDesde, String fechaHasta, String comercio,
			String directorio) throws Exception {

		Robot teclado = new Robot();
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);

		for (int i = 0; i < 21; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}

		DesgloceCaracteres(pan);

		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);

		for (int i = 0; i < 11; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}

		DesgloceCaracteres(fechaDesde);

		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);

		for (int i = 0; i < 11; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}

		DesgloceCaracteres(fechaHasta);

		for (int i = 0; i < 3; i++) {
			teclado.keyPress(KeyEvent.VK_TAB);
			teclado.keyRelease(KeyEvent.VK_TAB);
		}

		DesgloceCaracteres(comercio);
		screenshot(directorio);

		
		teclado.keyPress(KeyEvent.VK_F1);
		teclado.keyRelease(KeyEvent.VK_F1);
		Thread.sleep(2000);

		DesgloceCaracteres("d");
		screenshot(directorio);

		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);

		screenshot(directorio);
	}

	public static void obtenerDatos() throws Exception {

		BaseTest.driver.findElement(By.name("Copiar el texto seleccionado en el portapapeles")).click();
		String valor = copiarPortapapeles();
		String rutaProyecto = System.getProperty("user.dir");
		String ruta = rutaProyecto + "\\txt\\texto.txt";

		File file = new File(ruta);
		if (!file.exists()) {
			file.createNewFile();
		}
		FileWriter fw = new FileWriter(file);
		BufferedWriter bw = new BufferedWriter(fw);
		bw.write(valor);
		bw.close();

	}

	public static String fechaTrx() throws Exception {
		String fechaTrx = "";
		String rutaProyecto = System.getProperty("user.dir");
		File archivo = new File(rutaProyecto + "\\txt\\texto.txt");
		FileReader fr = new FileReader(archivo);
		BufferedReader br = new BufferedReader(fr);

		String linea;
		while ((linea = br.readLine()) != null) {
			if (linea.contains("Fecha Inicio:")) {
				fechaTrx = linea.substring(15, 25);
				fechaTrx = fechaTrx.replace("/", "");
				return fechaTrx;
			}
		

		}
		return fechaTrx;
	}

	public static String horaInicio() throws Exception {
		String horaInicio = "";
		String rutaProyecto = System.getProperty("user.dir");
		File archivo = new File(rutaProyecto + "\\txt\\texto.txt");
		FileReader fr = new FileReader(archivo);
		BufferedReader br = new BufferedReader(fr);

		String linea;
		while ((linea = br.readLine()) != null) {
			if (linea.contains("Fecha Inicio:")) {
				horaInicio = linea.substring(27, 35);
				horaInicio = horaInicio.replace(":", "");
				return horaInicio;
			}
	
		}
		return horaInicio;
	}

	public static String horaFin() throws Exception {
		String horaFinal = "";
		String rutaProyecto = System.getProperty("user.dir");
		File archivo = new File(rutaProyecto + "\\txt\\texto.txt");
		FileReader fr = new FileReader(archivo);
		BufferedReader br = new BufferedReader(fr);

		String linea;
		while ((linea = br.readLine()) != null) {
			if (linea.contains("Fecha Final.:")) {
				horaFinal = linea.substring(27, 35);
				horaFinal = horaFinal.replace(":", "");
				return horaFinal;
			}
		
		}
		return horaFinal;
	}

	public void pegarPortapapeles(String pegado) {
		StringSelection ss = new StringSelection(pegado);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(ss, null);
	}

	public static String copiarPortapapeles() throws UnsupportedFlavorException, IOException {
		Transferable t = Toolkit.getDefaultToolkit().getSystemClipboard().getContents(null);
		String copiado = (String) t.getTransferData(DataFlavor.stringFlavor);
		return copiado;
	}

	public static void ObtenerDataExcel() throws Exception {
		Robot teclado = new Robot();
		String[][] sheetData = null;
		String pan = "";
		String fechaDesde = "";
		String fechaHasta = "";
		String comercio = "";

		sheetData = BaseTest.retornaDatosExcel();
		for (int i = 1; i < sheetData.length; i++) {

			pan = sheetData[i][0].toString();
			fechaDesde = sheetData[i][1].toString();
			fechaHasta = sheetData[i][2].toString();
			comercio = sheetData[i][3].toString();
			
			System.out.println(pan);
			System.out.println(fechaDesde);
			System.out.println(fechaHasta);
			System.out.println(comercio);

			navegacionSwitch(pan, fechaDesde, fechaHasta, comercio, BaseTest.directorio);

			winium_page.obtenerDatos();

			String horaInicio = horaInicio();
			String horaFin = horaFin();
			String fechaTrx = fechaTrx();

			escribirCSV(fechaTrx, horaInicio, horaFin, comercio);

			for (int j = 0; j < 2; j++) {

				teclado.keyPress(KeyEvent.VK_F3);
				teclado.keyRelease(KeyEvent.VK_F3);
			}
		}
		

		

	}


	@SuppressWarnings("resource")
	public static void ObtenerDataExtraida(String directorio) throws Exception {

		String rutaProyecto = System.getProperty("user.dir");
		String ruta_csv = rutaProyecto +  "\\resources\\DataExtraida.csv";
		String fecha, horaInicio, horaFin, comercio;

        try{

    		Reader reader = Files.newBufferedReader(Paths.get(ruta_csv));
            CSVReader csvReader = new CSVReader(reader);
            String[] nextRecord;
             while ((nextRecord = csvReader.readNext()) != null) {
            	             	 
            	 fecha = nextRecord[0];
            	 horaInicio = nextRecord[1];
            	 horaFin = nextRecord[2];
            	 comercio = nextRecord[3];

            	 navergarCanal(fecha, horaInicio, horaFin, comercio, directorio);
 	 
             }

        }catch(Exception e){
        	e.getMessage().toString();
	
        }
	    
       
	}
	
	@SuppressWarnings("resource")
	public static void ObtenerDataExtraidaSatif(String directorio) throws Exception {

		String rutaProyecto = System.getProperty("user.dir");
		String ruta_csv = rutaProyecto +  "\\resources\\DataExtraida.csv";
		String fecha, horaInicio, horaFin, comercio;

        try{

    		Reader reader = Files.newBufferedReader(Paths.get(ruta_csv));
            CSVReader csvReader = new CSVReader(reader);
            String[] nextRecord;
             while ((nextRecord = csvReader.readNext()) != null) {
            	             	 
            	 fecha = nextRecord[0];
            	 horaInicio = nextRecord[1];
            	 horaFin = nextRecord[2];
            	 comercio = nextRecord[3];

            	 navergarSatif(fecha, horaInicio, horaFin, comercio, directorio);
 	 
             }

        }catch(Exception e){
        	e.getMessage().toString();
	
        }
	    
       
	}
	
	public static void navergarSatif(String fecha, String horaInicio, String horaFin, String comercio, String directorio) throws Exception {
		
		Robot teclado = new Robot();		

		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		for (int i = 0; i < 11; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}
		
		DesgloceCaracteres(fecha);
		
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		for (int i = 0; i < 9; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}
		
		DesgloceCaracteres(horaInicio);
		
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		for (int i = 0; i < 9; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}
		
		DesgloceCaracteres(horaFin);
		
		for (int i = 0; i < 2; i++) {
			teclado.keyPress(KeyEvent.VK_TAB);
			teclado.keyRelease(KeyEvent.VK_TAB);
		}
		
		DesgloceCaracteres("n");
		
		teclado.keyPress(KeyEvent.VK_F1);
		teclado.keyRelease(KeyEvent.VK_F1);
		
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		DesgloceCaracteres("i");
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
		teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
		teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
		teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		Thread.sleep(1000);
		
		teclado.keyPress(KeyEvent.VK_SHIFT);
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_SHIFT);
		
		DesgloceCaracteres("i");
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
		teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
		teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
		teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		teclado.keyPress(KeyEvent.VK_F11);
		teclado.keyRelease(KeyEvent.VK_F11);
		

		
	}
	
	 
	
	public static void navergarCanal(String fecha, String horaInicio, String horaFin, String comercio, String directorio) throws Exception {
			
		Robot teclado = new Robot();
		navegacionCanal(comercio);
		
		DesgloceCaracteres(fecha);
		
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		DesgloceCaracteres(horaInicio);
		
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		for (int i = 0; i < 9; i++) {
			teclado.keyPress(KeyEvent.VK_DELETE);
			teclado.keyRelease(KeyEvent.VK_DELETE);
		}
		
		DesgloceCaracteres(horaFin);
		
		for (int i = 0; i < 2; i++) {
			teclado.keyPress(KeyEvent.VK_TAB);
			teclado.keyRelease(KeyEvent.VK_TAB);
		}
		
		DesgloceCaracteres("n");
		
		teclado.keyPress(KeyEvent.VK_F1);
		teclado.keyRelease(KeyEvent.VK_F1);
		
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		
		DesgloceCaracteres("d");
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);

		teclado.keyPress(KeyEvent.VK_SHIFT);
		teclado.keyPress(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_TAB);
		teclado.keyRelease(KeyEvent.VK_SHIFT);
		
		DesgloceCaracteres("d");
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		screenshot(directorio);
		
		teclado.keyPress(KeyEvent.VK_ENTER);
		teclado.keyRelease(KeyEvent.VK_ENTER);
		
		
		for (int i = 0; i < 3; i++) {
			teclado.keyPress(KeyEvent.VK_F3);
			teclado.keyRelease(KeyEvent.VK_F3);
		}
		
		consultaCanal(directorio);
		
	}
		
		
	public static void escribirCSV(String fecha, String horaInicio, String horaFin, String comercio) {
		List<String[]> allData = new ArrayList<String[]>();
		String rutaProyecto = System.getProperty("user.dir");
		String ruta_csv = rutaProyecto +  "\\resources\\DataExtraida.csv";	
		File archivo =  new File(ruta_csv);
		
		
		try {
			if (!archivo.exists()) {
				
				CSVWriter crear = new CSVWriter(new FileWriter(ruta_csv));
//				String[] cabecera = new String [] {"Fecha","HoraInicio","HoraFin"};
//				crear.writeNext(cabecera);
				crear.close();
			}
			
			String [] datos = new String[] {fecha, horaInicio, horaFin, comercio};
			allData.add(datos);

			FileWriter fw = new FileWriter(ruta_csv, true);
			CSVWriter cw = new CSVWriter(fw);
			cw.writeAll(allData);	
			cw.close();
			fw.close();
			
		} catch (IOException e) {
			System.out.println("Error: "+e.getMessage());
		}
	}
	
	public static void navegacionCanal(String canal) throws Exception{
		
		Thread.sleep(1000);
		try {
			canal = canal.toUpperCase();
			
			Robot teclado = new Robot();
			switch (canal) {
			
			case "POS":
				
				for (int i = 0; i < 2; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 4; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "SOD":
				
				for (int i = 0; i < 3; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 4; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "TOT":
				
				for (int i = 0; i < 3; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 9; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "IM2":
				
				for (int i = 0; i < 1; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}

				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "RDF":
				
				for (int i = 0; i < 2; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 7; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "BDP":

				for (int i = 0; i < 2; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "SVP":
				
				for (int i = 0; i < 3; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 5; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "WEB":
				
				for (int i = 0; i < 4; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 2; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "ISW":
				
				for (int i = 0; i < 1; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 1; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
				
			case "QPY":
				
				for (int i = 0; i < 2; i++) {
					teclado.keyPress(KeyEvent.VK_PAGE_DOWN);
					teclado.keyRelease(KeyEvent.VK_PAGE_DOWN);
				}
				
				for (int i = 0; i < 5; i++) {
					teclado.keyPress(KeyEvent.VK_TAB);
					teclado.keyRelease(KeyEvent.VK_TAB);
				}
				
				teclado.keyPress(KeyEvent.VK_O);
				teclado.keyRelease(KeyEvent.VK_O);
				
				teclado.keyPress(KeyEvent.VK_ENTER);
				teclado.keyRelease(KeyEvent.VK_ENTER);
				
				DesgloceCaracteres("24");
				teclado.keyPress(KeyEvent.VK_F11);
				teclado.keyRelease(KeyEvent.VK_F11);
				
				teclado.keyPress(KeyEvent.VK_TAB);
				teclado.keyRelease(KeyEvent.VK_TAB);
				
				for (int i = 0; i < 11; i++) {
					teclado.keyPress(KeyEvent.VK_DELETE);
					teclado.keyRelease(KeyEvent.VK_DELETE);
				}
				
				break;
			default:
				break;
			}
			
		} catch (Exception e) {
			e.getMessage().toString();
		}
		
	}
	

	public static void DesgloceCaracteres(String codigoFinal) {
		try {

			Robot teclado = new Robot();
			for (int i = 0; i < codigoFinal.length(); i++) {

				char desglose = codigoFinal.charAt(i);
				String valor = Character.toString(desglose);

				switch (valor) {
				case "0":
					teclado.keyPress(KeyEvent.VK_0);
					teclado.keyRelease(KeyEvent.VK_0);
					break;
				case "1":
					teclado.keyPress(KeyEvent.VK_1);
					teclado.keyRelease(KeyEvent.VK_1);
					break;
				case "2":
					teclado.keyPress(KeyEvent.VK_2);
					teclado.keyRelease(KeyEvent.VK_2);
					break;
				case "3":
					teclado.keyPress(KeyEvent.VK_3);
					teclado.keyRelease(KeyEvent.VK_3);
					break;
				case "4":
					teclado.keyPress(KeyEvent.VK_4);
					teclado.keyRelease(KeyEvent.VK_4);
					break;
				case "5":
					teclado.keyPress(KeyEvent.VK_5);
					teclado.keyRelease(KeyEvent.VK_5);
					break;
				case "6":
					teclado.keyPress(KeyEvent.VK_6);
					teclado.keyRelease(KeyEvent.VK_6);
					break;
				case "7":
					teclado.keyPress(KeyEvent.VK_7);
					teclado.keyRelease(KeyEvent.VK_7);
					break;
				case "8":
					teclado.keyPress(KeyEvent.VK_8);
					teclado.keyRelease(KeyEvent.VK_8);
					break;
				case "9":
					teclado.keyPress(KeyEvent.VK_9);
					teclado.keyRelease(KeyEvent.VK_9);
					break;
				case "a":
					teclado.keyPress(KeyEvent.VK_A);
					teclado.keyRelease(KeyEvent.VK_A);
					break;
				case "b":
					teclado.keyPress(KeyEvent.VK_B);
					teclado.keyRelease(KeyEvent.VK_B);
					break;
				case "c":
					teclado.keyPress(KeyEvent.VK_C);
					teclado.keyRelease(KeyEvent.VK_C);
					break;
				case "d":
					teclado.keyPress(KeyEvent.VK_D);
					teclado.keyRelease(KeyEvent.VK_D);
					break;
				case "e":
					teclado.keyPress(KeyEvent.VK_E);
					teclado.keyRelease(KeyEvent.VK_E);
					break;
				case "f":
					teclado.keyPress(KeyEvent.VK_F);
					teclado.keyRelease(KeyEvent.VK_F);
					break;
				case "g":
					teclado.keyPress(KeyEvent.VK_G);
					teclado.keyRelease(KeyEvent.VK_G);
					break;
				case "h":
					teclado.keyPress(KeyEvent.VK_H);
					teclado.keyRelease(KeyEvent.VK_H);
					break;
				case "i":
					teclado.keyPress(KeyEvent.VK_I);
					teclado.keyRelease(KeyEvent.VK_I);
					break;
				case "j":
					teclado.keyPress(KeyEvent.VK_J);
					teclado.keyRelease(KeyEvent.VK_J);
					break;
				case "k":
					teclado.keyPress(KeyEvent.VK_K);
					teclado.keyRelease(KeyEvent.VK_K);
					break;
				case "l":
					teclado.keyPress(KeyEvent.VK_L);
					teclado.keyRelease(KeyEvent.VK_L);
					break;
				case "m":
					teclado.keyPress(KeyEvent.VK_M);
					teclado.keyRelease(KeyEvent.VK_M);
					break;
				case "n":
					teclado.keyPress(KeyEvent.VK_N);
					teclado.keyRelease(KeyEvent.VK_N);
					break;
				case "o":
					teclado.keyPress(KeyEvent.VK_O);
					teclado.keyRelease(KeyEvent.VK_O);
					break;
				case "p":
					teclado.keyPress(KeyEvent.VK_P);
					teclado.keyRelease(KeyEvent.VK_P);
					break;
				case "q":
					teclado.keyPress(KeyEvent.VK_Q);
					teclado.keyRelease(KeyEvent.VK_Q);
					break;
				case "r":
					teclado.keyPress(KeyEvent.VK_R);
					teclado.keyRelease(KeyEvent.VK_R);
					break;
				case "s":
					teclado.keyPress(KeyEvent.VK_S);
					teclado.keyRelease(KeyEvent.VK_S);
					break;
				case "t":
					teclado.keyPress(KeyEvent.VK_T);
					teclado.keyRelease(KeyEvent.VK_T);
					break;
				case "u":
					teclado.keyPress(KeyEvent.VK_U);
					teclado.keyRelease(KeyEvent.VK_U);
					break;
				case "v":
					teclado.keyPress(KeyEvent.VK_V);
					teclado.keyRelease(KeyEvent.VK_V);
					break;
				case "w":
					teclado.keyPress(KeyEvent.VK_W);
					teclado.keyRelease(KeyEvent.VK_W);
					break;
				case "x":
					teclado.keyPress(KeyEvent.VK_X);
					teclado.keyRelease(KeyEvent.VK_X);
					break;
				case "y":
					teclado.keyPress(KeyEvent.VK_Y);
					teclado.keyRelease(KeyEvent.VK_Y);
					break;
				case "z":
					teclado.keyPress(KeyEvent.VK_Z);
					teclado.keyRelease(KeyEvent.VK_Z);
					break;
				default:

					break;
				}
			}
		} catch (Exception e) {
			System.out.println("Error: " + e.getMessage().toString());
		}
	}

	public static void screenshot(String directorioHora) throws IOException {
		try {
			Thread.sleep(1500);
			Date objDate = new Date();
			String FechaFormateada;
			String strDateFormat = "ddMMyyyy hhmmss";
			SimpleDateFormat objSDF = new SimpleDateFormat(strDateFormat);
			FechaFormateada = objSDF.format(objDate);
			File srcFile = BaseTest.driver.getScreenshotAs(OutputType.FILE);
			FechaFormateada = directorioHora + "/" + FechaFormateada + ".png";
			File targetFile = new File(FechaFormateada);
			FileUtils.copyFile(srcFile, targetFile);

		} catch (Exception e) {
			System.out.println(e.getMessage().toString());

		}
	}
	
	public static void generarPDF(String ruta) {
		String rutaProyecto = System.getProperty("user.dir");
		ruta = ruta + "/";
		String causa = null;
		try {
			
			// Creacion de Documento PDF
			Document documento = new Document();
			// Asignacion de Nombre y extencion al PDF
			FileOutputStream ficheroPdf = new FileOutputStream("Evidencia.pdf");
			// Escritura separada por 20MM
			PdfWriter.getInstance(documento, ficheroPdf).setInitialLeading(20);

			// Abre el Documento
			documento.open();
			Calendar cal = Calendar.getInstance();
			Date fecha = new Date(cal.getTimeInMillis());
			SimpleDateFormat formato = new SimpleDateFormat("dd/MM/yyyy");
			String FechaFormateada = formato.format(fecha);
			Paragraph anchor = new Paragraph("Evidencia de Ejecución Evidencia Switch");
			anchor.setAlignment(Element.ALIGN_LEFT);
			Paragraph fechaPDF = new Paragraph("Fecha de Ejecución: " + FechaFormateada);

			// Agregar Texto al PDF
			documento.add(anchor);
			documento.add(Chunk.NEWLINE);
			documento.add(fechaPDF);
			documento.add(Chunk.NEWLINE);

			// Agregar las Imagenes Capturadas en la Ejecucion

			String[] files = getFiles(ruta);

			if (files != null) {

				int size = files.length;

				for (int i = 0; i < size; i++) {
					Image foto = Image.getInstance(files[i]);
					foto.scaleToFit(500, 500);
					foto.setAlignment(Chunk.ALIGN_MIDDLE);
					documento.add(foto);
				}
			}
			// cierra el documento
			documento.close();

			// Copia el documento de la ruta que se crea a la ruta de evidencia de ejecucion
			
			
			Path origenPath = FileSystems.getDefault()
					.getPath(rutaProyecto+ "\\Evidencia.pdf");
			Path destinoPath = FileSystems.getDefault().getPath(ruta + "Evidencia.pdf");

			try {
				Files.move(origenPath, destinoPath, StandardCopyOption.REPLACE_EXISTING);
			} catch (IOException e) {
				System.err.println(e);
			}

		} catch (Exception e) {
			System.out.println(e.getMessage().toString());

		}

	}

	public static String[] getFiles(String dir_path) {

		String[] arr_res = null;

		File f = new File(dir_path);

		if (f.isDirectory()) {
			// Obtiene el listado de archivos que hay en el directorio
			File[] arr_content = f.listFiles();
			int size = arr_content.length;
			List<String> res = new ArrayList<>();

			// recorre los archivos en el directorio y los agrega al array
			for (int i = 0; i < size; i++) {
				if (arr_content[i].isFile())
					res.add(arr_content[i].toString());
			}
			// Ordena el Array de menor a mayor
			Collections.sort(res);

			if (res.size() > 0)
				arr_res = res.toArray(new String[0]);

		} else

			System.err.println("¡ Ruta NO válida !");

		return arr_res;
	}
	
}
