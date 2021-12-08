package com.carl.excel.config;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

/**
 * @author: He Dong
 */
public class ConfigXlsx2Xml {
	
	private String xmlPath;
	private String xlsxPath;
	private String sheetName;
	private String sex;
	private String name;

	public String getXinlao() {
		return xinlao;
	}

	public void setXinlao(String xinlao) {
		this.xinlao = xinlao;
	}

	public String getBumen() {
		return bumen;
	}

	public void setBumen(String bumen) {
		this.bumen = bumen;
	}

	private String xinlao;
	private String bumen;

	public String getXmlPath() {
		return xmlPath;
	}
	public void setXmlPath(String path) {
		this.xmlPath = path;
	}
	public String getXlsxPath() {
		return xlsxPath;
	}
	public void setXlsxPath(String xlsxPath) {
		this.xlsxPath = xlsxPath;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public String getSex() {
		return sex;
	}
	public void setSex(String value) {
		this.sex = value;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append("--------------CONFIG-------------")
		.append("xmlPath="+ xmlPath)
		.append("\nxlsxPath=" + xlsxPath)
		.append("\nsheetName=" + sheetName)
		.append("\nvalue=" + sex)
		.append("\nname=" + name)
		.append("\n--------------END-------------");
		return sb.toString();
	}
	
	public static List<ConfigXlsx2Xml> parse(String configPath) {
		File f = new File(configPath);
		System.out.println("f.exists()"+f.exists());
		if (!f.exists()) {
			return null;
		}
		List<ConfigXlsx2Xml> list = new ArrayList<ConfigXlsx2Xml>();
		BufferedReader br = null;
		try {
			br = new BufferedReader(new InputStreamReader(new FileInputStream(configPath), "UTF-8"));
			String line = null;
			ConfigXlsx2Xml config = null;
			while((line = br.readLine()) != null) {
				line = line.trim();
				
				if (line.startsWith("<xlsx2xml>")) {
					config = new ConfigXlsx2Xml();
				} else if (line.startsWith("</xlsx2xml>")) {
					if (config != null) {
						list.add(config);
					}
				} else {
					if (config == null) {
						continue;
					}
					String[] array = line.split("=");
					if (array.length != 2) {
						continue;
					}
					array[0] = array[0].trim();
					if (array[0].equalsIgnoreCase("xmlPath")) {
						config.setXmlPath(array[1]);
					} else if (array[0].equalsIgnoreCase("xlsxPath")) {
						config.setXlsxPath(array[1]);
					} else if (array[0].equalsIgnoreCase("sheetName")) {
						config.setSheetName(array[1]);
					} else if (array[0].equalsIgnoreCase("sex")) {
						config.setSex(array[1]);
					}  else if (array[0].equalsIgnoreCase("name")) {
						config.setName(array[1]);
					} else if (array[0].equalsIgnoreCase("xinlao")) {
						config.setXinlao(array[1]);
					} else if (array[0].equalsIgnoreCase("bumen")) {
						config.setBumen(array[1]);
					}
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				br.close();
				br = null;
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return list;
	}
}
