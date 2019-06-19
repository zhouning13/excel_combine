package com.rubyren.savedrecs;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadTxtAndReplaceExcel {

	public static void main(String[] args) throws IOException {
		// 数据收集
		Path p = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/下载补充数据20190403/");

		Path[] files = Files.walk(p).sorted().toArray(Path[]::new);
		Map<String, String[]> contents = new HashMap<>();
		for (Path f : files) {
			if (Files.isDirectory(f)) {
				continue;
			}
			if (".DS_Store".equalsIgnoreCase(f.getFileName().toString())) {
				continue;
			}
			List<String> lines = Files.readAllLines(f);
			lines.remove(0);
			for (String line : lines) {
				String[] fs = line.split("\t");
				contents.put(fs[8], fs);
			}

		}

		// excel替换
		Path from = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/WOS 24711条数据原始.xlsx");
		Path to = Paths
				.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/WOS 24711条数据_20190407.xlsx");
		try (FileInputStream stream = new FileInputStream(from.toFile());
				XSSFWorkbook workbook = new XSSFWorkbook(stream);
				FileOutputStream out = new FileOutputStream(to.toFile())) {
			
			Font font = workbook.createFont();
			// font.setFontName("Arial");
			font.setFontName("等线");
			font.setFontHeight((short) 220);
			font.setBold(false);
			// 把字体应用到当前的样式
			
			CellStyle errorStyle = workbook.createCellStyle();
			errorStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
			errorStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			errorStyle.setFillPattern(FillPatternType.ALT_BARS);
			errorStyle.setFont(font);
			
			XSSFSheet sheet = workbook.getSheetAt(0);
			int i = 0;
			for (Row row : sheet) {
				String cr = get(row, 29);
				if (StringUtils.isBlank(cr)) {
					String ti = get(row, 8);
					if (ti.startsWith("=")) {
						ti = ti.replaceFirst("=", "");
					}
					String[] content = contents.remove(ti);
					if (content == null) {
						System.out.println("未找到数据 " + i + " 行， ti 为 " + ti);
					} else {
						System.out.println("匹配到数据 " + i + " 行， ti 为 " + ti);
						for (int j = 0; j <= 67; j++) {
							if (StringUtils.isNotBlank(content[j])) {
								try {
									Cell cell = row.getCell(j);
									if (cell == null) {
										cell = row.createCell(j);
									}
									cell.setCellStyle(errorStyle);
									cell.setCellValue(content[j]);
								} catch (IllegalArgumentException e) {
									// DONOTHING
								} catch (Exception e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							} else {
								//System.out.println();
							}
						}
					}
				}
				i++;
			}

			for (String ti : contents.keySet()) {
				System.out.println("未匹配到题名 ti 为 " + ti);
			}
			workbook.write(out);
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
}
