package utilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class PropertyFileUtil {
	public static String getValueFoeKey(String Key) throws IOException{
		Properties p=new Properties();
		FileInputStream fis=new FileInputStream("D:\\RAJIB\\Selenium\\StockAccounting_Hybrid_mvn82\\PropertiesFile\\Enviroment.properties");
		p.load(fis);
		return p.getProperty(Key);
		
	}

}
