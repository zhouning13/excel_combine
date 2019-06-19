package com.rubyren.wos;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashSet;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FilteExcel {
	public static void main(String[] args) {
		Path filter = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/wos/WOS-24711.xlsx");
		Path from = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/wos/WOS TI 24756.xlsx");
		Path to = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/wos/WOS_compared.xlsx");

		HashSet<String> existedName = new HashSet<>();
		HashSet<String> existedPair = new HashSet<>();
		try (FileInputStream in = new FileInputStream(filter.toFile());
				XSSFWorkbook fromWorkbook = new XSSFWorkbook(in);) {
			XSSFSheet sheet = fromWorkbook.getSheetAt(0);
			int line = 0;
			for (Row row : sheet) {
				if (line != 0) {
					String name = row.getCell(8).getStringCellValue().trim();
					String doi = "";
					if (row.getCell(54) != null) {
						doi = row.getCell(54).getStringCellValue().trim();
					}

					if (existedPair.contains(name + " | " + doi)) {
						System.out.println(name + " | " + doi);
					} else {
						existedPair.add(name + " | " + doi);
					}

					if (existedName.contains(name)) {
						System.out.println(name);
					} else {
						existedName.add(name);
					}
				}
				line++;
			}
			System.out.println(line + " : " + existedPair.size());
			System.out.println();

		} catch (Exception e) {
			e.printStackTrace();
		}

		try (FileInputStream in = new FileInputStream(from.toFile());
				XSSFWorkbook fromWorkbook = new XSSFWorkbook(in);
				XSSFWorkbook toWorkbook = new XSSFWorkbook();
				FileOutputStream out = new FileOutputStream(to.toFile())) {
			XSSFSheet fromSheet = fromWorkbook.getSheetAt(0);
			XSSFSheet toSheet = toWorkbook.createSheet();

			CellStyle style = toWorkbook.createCellStyle();
			Font font = toWorkbook.createFont();
			// font.setFontName("Arial");
			font.setFontName("等线");
			font.setFontHeight((short) 220);
			font.setBold(false);
			font.setColor(HSSFColorPredefined.RED.getIndex());
			// 把字体应用到当前的样式
			style.setFillBackgroundColor(HSSFColorPredefined.RED.getIndex());
			style.setFont(font);

			int line = 0;
			int lineTo = 0;
			for (Row row : fromSheet) {
				if (line != 0) {
					String name = row.getCell(1).getStringCellValue().trim();
					String doi = "";
					if (row.getCell(5) != null) {
						doi = row.getCell(5).getStringCellValue().trim();
					}

					System.out.println(name + " | " + doi);
					if (!existedPair.contains(name + " | " + doi)) {
						Row toRow = toSheet.createRow(lineTo);
						

						for (int i = 0; i < 9; i++) {
							copy(row, i, toRow, i);
						}
						if (!existedName.contains(name)) {
							toRow.createCell(9).setCellValue("1");
						}
						lineTo++;
					}
				}
				line++;
			}
			toWorkbook.write(out);
			out.flush();
		} catch (Exception e) {
			e.printStackTrace();
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
