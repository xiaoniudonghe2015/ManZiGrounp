package com.carl.excel.config;
/**
 * @author: He Dong
 */
public class ConfigUtil {
	
	public static String PATH_CONFIG_XML2XLSX;
	public static String PATH_CONFIG_XLSX2XML;
	private static String sCurDir;
	
	static {
		sCurDir = System.getProperty("user.dir");///home/dh/aa/ConvExcel2Xml
		PATH_CONFIG_XML2XLSX = sCurDir + "/app/config/XML2XLSX.config";
		PATH_CONFIG_XLSX2XML = sCurDir + "/app/config/XLSX2XML.config";
	}
}
