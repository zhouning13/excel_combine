package com.rubyren.quota.service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndWriteDataFile {
	private static Path dataFile = Paths.get("/Users/zhouning/Desktop/麦达数据-任国华梳理.xlsx");
	private static Path outFile = Paths.get("/Users/zhouning/Desktop/麦达数据-out.xlsx");

	public static void read() {

		Map<String, String[]> eduDatas = ReadEduGvnFile.read();
		try (FileInputStream stream = new FileInputStream(dataFile.toFile());
				XSSFWorkbook workbook = new XSSFWorkbook(stream);
				FileOutputStream out = new FileOutputStream(outFile.toFile())) {
			XSSFSheet sheet = workbook.getSheetAt(1);

			for (Row row : sheet) {
				String name = get(row, 3);
				String[] eduData = eduDatas.get(name);
				if (eduData != null) {
					set(row, 8, "是");
					set(row, 9, eduData[0]);
					set(row, 10, eduData[1]);
					set(row, 11, eduData[2]);
				}

			}
			workbook.write(out);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static String get(Row row, int i) {
		Cell c = row.getCell(i);
		if (c == null) {
			return null;
		}
		switch (c.getCellTypeEnum()) {
		case BLANK:
			return "";
		case NUMERIC:
			Double d = new Double(c.getNumericCellValue());
			return d.intValue() + "";
		default:
			return c.getStringCellValue();
		}
	}

	private static void set(Row row, int i, String value) {
		Cell c = row.getCell(i);
		if (c == null) {
			return;
		}
		c.setCellValue(value);
	}

	public static void main(String[] args) {
		read();
	}
}
