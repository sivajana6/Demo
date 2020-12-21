package com.document;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSReader {

	public static void main(String args[]) throws Exception {
		List<String> headerList = new ArrayList<String>();
		Map<String, String> valueMap = new TreeMap<String, String>();
		Map<Integer, Map<String, String>> finalMap = new TreeMap<Integer, Map<String, String>>();

		File file = new File("D:/Data.xls");
		String extension = file.getName().substring(file.getName().indexOf("."));
		FileInputStream fis = new FileInputStream(file);
		Workbook workBook = null;
		if(extension.equalsIgnoreCase(".xls")) {
			workBook = new HSSFWorkbook(fis);
		} else if(extension.equalsIgnoreCase(".xlsx")) {
			workBook = new XSSFWorkbook(fis);
		}
		
		List<PictureData> pictures = (List<PictureData>)workBook.getAllPictures();
		Iterator<PictureData> iterator = pictures.iterator();
		while(iterator.hasNext()) {
			PictureData picData = iterator.next();
			byte[] data = picData.getData();
			if(picData.suggestFileExtension().equals("jpg")) {
				FileOutputStream fos = new FileOutputStream("XLSpicture.jpg");
				fos.write(data);
				fos.close();
			}
		}
		
		Sheet sheet = workBook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum();
		for(int i = 0; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			int cols = row.getPhysicalNumberOfCells();
			for(int j = 0; j < cols; j++) {
				if(i == 0) {
					headerList.add(row.getCell(j).getStringCellValue());
				} else {
					valueMap.put(headerList.get(j), row.getCell(j).getStringCellValue());
				}
			}
			if(i != 0) {
				finalMap.put(i, valueMap);
			}
		}
		System.out.println(finalMap);
		workBook.close();
	}
}