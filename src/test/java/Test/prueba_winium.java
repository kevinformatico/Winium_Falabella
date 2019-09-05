package Test;

import Base.BaseTest;
import Page.winium_page;

public class prueba_winium extends BaseTest {
	public prueba_winium() {

		super();
	}

	public static void main(String[] args) throws Exception {
		BaseTest.crearDirectorios();
		BaseTest.abrirSwitch();
		winium_page.loginSwitch(leerMatrizTxt("usuario", matriz) ,leerMatrizTxt("clave", matriz), directorio);
		winium_page.consultaTrx(directorio);
		winium_page.ObtenerDataExcel();
		BaseTest.cerrarApp();
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
		winium_page.generarPDF(directorio);
		
	}

}

