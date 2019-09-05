package Page;

import java.net.MalformedURLException;
import java.net.URL;

import org.openqa.selenium.By;
import org.openqa.selenium.winium.DesktopOptions;
import org.openqa.selenium.winium.WiniumDriver;

import Base.BaseTest;

public class tn5250_page {
	
public static void abrirTN5250() throws MalformedURLException, Exception {
		
		BaseTest.directorio = BaseTest.crearDirectorio();
		
		DesktopOptions option = new DesktopOptions();
		option.setApplicationPath("C:\\Program Files (x86)\\TN5250\\TN5250.EXE");
		
		BaseTest.driver = new WiniumDriver(new URL("http://localhost:9999"), option);
		BaseTest.matriz =  BaseTest.getDataTxt("DataSwitch");
	}

	
public static void abrirAplicacion(String directorio) throws Exception {
		
		BaseTest.driver.findElement(By.name("Host to connect to:")).sendKeys("10.140.0.215");
		BaseTest.driver.findElement(By.name("OK")).click();

	}


}
