package com.carl.excel;

import com.example.fenzhuo.Staff;
import com.example.fenzhuo.Table;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author: He Dong
 */
public class ExcelUtil {
	
	private static final String KEY_NAME = "name";

	public static void writeExcel(String path, String sheetName,
								  ArrayList<Table> list, int maxCount)
			throws InvalidFormatException, IOException {
		if (list == null || list.size() <= 0) {
			return;
		}
		File f = new File(path);
		boolean newFile = !f.exists();
		XSSFWorkbook xssfWorkbook = null;
		FileInputStream fis = null;
		if (newFile) {
			xssfWorkbook = new XSSFWorkbook();
		} else {
			fis = new FileInputStream(f);
			xssfWorkbook = new XSSFWorkbook(fis);
		}

		if (xssfWorkbook.getSheet(sheetName) != null) {
			xssfWorkbook.removeSheetAt(xssfWorkbook.getSheetIndex(sheetName));
		}
		xssfWorkbook.createSheet(sheetName);
		XSSFSheet xssfSheet = xssfWorkbook.getSheet(sheetName);
		for (int i = 0; i < maxCount + 1; i++) {
			xssfSheet.createRow(i);
		}
		for (int i = 0; i < list.size(); i++) {
			Table table = list.get(i);
			XSSFCell cell = xssfSheet.getRow(0).createCell(i);
			cell.setCellValue(table.tableId);
			for (int j = 1; j < table.staffArrayList.size() + 1; j++) {
				cell = xssfSheet.getRow(j).createCell(i);
				Staff staff = table.staffArrayList.get(j - 1);
				cell.setCellValue(staff.name + "(" + staff.sex + "," + staff.status + "," + staff.section + ")");
			}
		}
		FileOutputStream fos = new FileOutputStream(path, false);
		fos.flush();
		xssfWorkbook.write(fos);
		if (fis != null) {
			fis.close();
		}
		fos.close();

		xssfWorkbook.close();
	}

	
	public static ArrayList<Staff> ReadExcel(String path, String sheetName,
										String name, String sex, String bumen, String xinlao)
			throws InvalidFormatException, IOException {
		if (sheetName == null || sex == null || name == null) {
			return null;
		}
		File f = new File(path);
		if (!f.exists()) {
			return null;
		}
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(f);
		System.out.println("sheetName"+sheetName);
		XSSFSheet xssfSheet = xssfWorkbook.getSheet(sheetName);
		XSSFRow row0 = xssfSheet.getRow(0);
		int cellEnd = row0.getLastCellNum();
		int indexSex = -1;
		int indexName = -1;
		int indexXinlao = -1;
		int indexBumen = -1;
		System.out.println("cellEnd-->"+cellEnd);
		for (int i = 0; i <= cellEnd; i ++) {
			XSSFCell xssfCell = row0.getCell(i);
			if (xssfCell != null && sex.equals(xssfCell.getStringCellValue())) {
				indexSex = i;
				break;
			}
		}
		for (int i = 0; i <= cellEnd; i ++) {
			XSSFCell xssfCell = row0.getCell(i);
			if (xssfCell != null && name.equals(xssfCell.getStringCellValue())) {
				indexName = i;
				break;
			}
		}
		for (int i = 0; i <= cellEnd; i ++) {
			XSSFCell xssfCell = row0.getCell(i);
			if (xssfCell != null && xinlao.equals(xssfCell.getStringCellValue())) {
				indexXinlao = i;
				break;
			}
		}
		for (int i = 0; i <= cellEnd; i ++) {
			XSSFCell xssfCell = row0.getCell(i);
			if (xssfCell != null && bumen.equals(xssfCell.getStringCellValue())) {
				indexBumen = i;
				break;
			}
		}
		if (indexSex < 0 || indexName < 0) {
			return null;
		}
		ArrayList<Staff> list = new ArrayList<Staff>();
		int rowEnd = xssfSheet.getLastRowNum();
		System.out.println("rowEnd-->"+rowEnd);
		for (int i = 1; i <= rowEnd; i ++) {
			XSSFRow row = xssfSheet.getRow(i);
			if (row == null) {
				continue;
			}
			XSSFCell cellName = row.getCell(indexName);
			if (cellName == null) {
				continue;
			}
			XSSFCell cellSex = row.getCell(indexSex);
			if (cellSex == null) {
				continue;
			}
			XSSFCell cellXinlao = row.getCell(indexXinlao);
			if (cellXinlao == null) {
				continue;
			}
			XSSFCell cellBumen = row.getCell(indexBumen);
			if (cellBumen == null) {
				continue;
			}
			Staff entry = new Staff();
			entry.name = cellName.getStringCellValue();
			entry.sex = cellSex.getStringCellValue();
			entry.status = cellXinlao.getStringCellValue();
			entry.section = cellBumen.getStringCellValue();
			list.add(entry);
			
		}
		xssfWorkbook.close();
		return list;
	}

}
