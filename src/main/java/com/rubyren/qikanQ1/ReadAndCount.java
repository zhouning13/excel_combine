package com.rubyren.qikanQ1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndCount {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Path from = Paths.get("/Users/zhouning/Documents/ruby/Q1期刊处理版.xlsx");
		Path to = Paths.get("/Users/zhouning/Documents/ruby/Q1期刊处理版_20190401.xlsx");

		try (FileInputStream in = new FileInputStream(from.toFile());
				XSSFWorkbook fromWorkbook = new XSSFWorkbook(in);
				XSSFWorkbook toWorkbook = new XSSFWorkbook();
				FileOutputStream out = new FileOutputStream(to.toFile())) {

			XSSFSheet sheet2 = fromWorkbook.getSheetAt(1);
			HashMap<String, Integer> issns = new HashMap<>();
			HashMap<String, Integer> names = new HashMap<>();
			HashMap<Integer, String[]> issnNameRevl = new HashMap<>();
			TreeMap<Integer, Integer> lineMap = new TreeMap<>();
			int line = 0;
			for (Row row : sheet2) {
				if (line != 0) {
					String name = StringUtils.trimToEmpty(get(row, 0));
					String issn = StringUtils.trimToEmpty(get(row, 1));
					names.put(name, line);
					issns.put(issn, line);
					issnNameRevl.put(line, new String[] { name, issn });
				}
				line++;
			}
			System.out.println(issns.size());

			XSSFSheet sheet1 = fromWorkbook.getSheetAt(0);
			XSSFSheet sheetTo1 = toWorkbook.createSheet();

			CellStyle style = toWorkbook.createCellStyle();
			Font font = toWorkbook.createFont();
			// font.setFontName("Arial");
			font.setFontName("等线");
			font.setFontHeight((short) 220);
			font.setBold(false);
			// 把字体应用到当前的样式
			style.setFont(font);
			int index = 0, indexTo = 1;
			sheetTo1.createRow(0);
			for (Row row : sheet1) {

				if (index > 1) {
					String name = StringUtils.trimToEmpty(get(row, 3));
					String issn = StringUtils.trimToEmpty(get(row, 12));
					int i = 0;
					if (issns.containsKey(issn)) {
						i = issns.get(issn);
					} else if (names.containsKey(name)) {
						i = names.get(name);
					}
					if (i != 0) {
						Integer j = lineMap.get(i);
						if (j != null) {
							lineMap.put(i, j.intValue() + 1);
						} else {
							lineMap.put(i, 1);
						}

						Row rowTo = sheetTo1.createRow(indexTo);
						copy(row, 2, rowTo, 0);
						copy(row, 3, rowTo, 1);
						copy(row, 4, rowTo, 2);
						copy(row, 10, rowTo, 3);
						copy(row, 11, rowTo, 4);
						copy(row, 12, rowTo, 5);
						copy(row, 13, rowTo, 6);
						copy(row, 14, rowTo, 7);
						copy(row, 15, rowTo, 8);
						copy(row, 16, rowTo, 9);
						copy(row, 17, rowTo, 10);
						Cell c11 = rowTo.createCell(11);
						c11.setCellValue(i);
						indexTo++;
					}
				}
				index++;
			}
			XSSFSheet sheetTo2 = toWorkbook.createSheet();
			sheetTo2.createRow(0);
			int k = 1;
			for (Map.Entry<Integer, Integer> entry : lineMap.entrySet()) {
				Row rowTo = sheetTo2.createRow(k);
				int lineNum = entry.getKey();
				Cell c0 = rowTo.createCell(0);
				c0.setCellValue(lineNum);

				String[]pair =issnNameRevl.get(lineNum);
				Cell c1 = rowTo.createCell(1);
				c1.setCellValue(pair[0]);
				
				Cell c2 = rowTo.createCell(2);
				c2.setCellValue(pair[1]);
				
				Cell c3 = rowTo.createCell(3);
				c3.setCellValue(entry.getValue());
				k++;
			}
			toWorkbook.write(out);
			out.flush();
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

	private static void copy(Row fromRow, int fromIndex, Row toRow, int toIndex) {
		Cell c = fromRow.getCell(fromIndex);
		if (c == null) {
			return;
		}
		Cell to = toRow.createCell(toIndex);
		switch (c.getCellTypeEnum()) {
		case BLANK:
			return;
		case NUMERIC:
			to.setCellValue(c.getNumericCellValue());
			return;
		default:
			to.setCellValue(c.getStringCellValue());
		}
	}
}
