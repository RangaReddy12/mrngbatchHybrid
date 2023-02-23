package utilities;

import java.io.FileInputStream;
import java.util.Properties;

public class PropertyFileUtil {
public static Properties conprop;
public static String getValueForKey(String key)throws Throwable
{
	conprop = new Properties();
	conprop.load(new FileInputStream("D:\\MrngBatch_OJT\\StockAccounting_ERP\\PropertyFiles\\Environment.properties"));
	return conprop.getProperty(key);
}
}
