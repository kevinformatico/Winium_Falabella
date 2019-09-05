package Test;

import Base.BaseTest;
import Page.tn5250_page;

public class PruebaTN5250 extends BaseTest{
	public PruebaTN5250() {

		super();
	}
	public static void main(String[] args) throws Exception {
		
	tn5250_page.abrirTN5250();
	tn5250_page.abrirAplicacion(directorio);

	}
}