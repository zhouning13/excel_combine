package com.rubyren.quota.service;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadEduGvnFile {
	private static Path eduFile = Paths.get("/Users/zhouning/Desktop/教育部指标.xlsx");

	public static Map<String, String[]> read() {
		Map<String, String[]> result = new HashMap<>();
		try (FileInputStream stream = new FileInputStream(eduFile.toFile());
				XSSFWorkbook workbook = new XSSFWorkbook(stream)) {
			XSSFSheet sheet = workbook.getSheetAt(1);

			for (Row row : sheet) {
				String name = get(row, 3);
				String year = get(row, 6);
				String amount = get(row, 7);
				String url = get(row, 12);
				if (StringUtils.isNotBlank(name)) {
					result.put(name, new String[] { year, amount, url });
				}
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println(result);
		return result;
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

	public static void main(String[] args) {
		read();
	}
}
