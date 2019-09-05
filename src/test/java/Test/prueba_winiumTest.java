package Test;

import org.testng.annotations.Test;

import Base.BaseTest;
import Page.winium_page;

public class prueba_winiumTest extends BaseTest {
	public prueba_winiumTest() {

		super();
	}
	
	@Test
	public void test() {
		try {
			BaseTest.crearDirectorios();
			BaseTest.abrirSwitch();
			winium_page.loginSwitch(leerMatrizTxt("usuario", matriz) ,leerMatrizTxt("clave", matriz), directorio);
			winium_page.consultaTrx(directorio);
			winium_page.ObtenerDataExcel();
			BaseTest.cerrarApp();
			winium_page.generarPDF(directorio);
		} catch (Exception e) {
			System.out.println("Error: "+e.getMessage());
			e.printStackTrace();
		}

//		BaseTest.abrirSwitch();
//		winium_page.loginSwitch(leerMatrizTxt("usuario", matriz), leerMatrizTxt("clave", matriz), directorio);
//		winium_page.consultaCanal(directorio);
//		winium_page.ObtenerDataExtraida(directorio);
//		BaseTest.cerrarApp();
//		BaseTest.abrirSwitch();
//		winium_page.loginSwitch(leerMatrizTxt("usuario", matriz) ,leerMatrizTxt("clave", matriz), directorio);
//		winium_page.consultaSatif(directorio);
//		winium_page.ObtenerDataExtraidaSatif(directorio);
//		BaseTest.cerrarApp();
		
		
	}

}

