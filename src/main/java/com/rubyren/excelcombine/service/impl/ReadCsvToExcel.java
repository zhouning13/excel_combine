package com.rubyren.excelcombine.service.impl;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import com.rubyren.excelcombine.model.Database;
import com.rubyren.excelcombine.model.Institution;

public class ReadCsvToExcel {
	NamingServiceImpl namingService = new NamingServiceImpl();

	public void readCsvAndTurn() throws IOException {
		Path from = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/趋势数据201810/");
		Path to = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/趋势201810to/");
		Files.walk(from, 1).sorted().forEach(f -> {
			if (!Objects.equals(from, f)) {

				try {
					String name = f.getFileName().toString();
					name = name.substring(0, name.length() - 4);
					System.out.println(name);
					Workbook workbook = new SXSSFWorkbook();
					FileOutputStream out = new FileOutputStream(to.resolve(name + ".xlsx").toFile());
					Sheet sheet = workbook.createSheet(name);
					sheet.setDefaultColumnWidth(9);
					sheet.setDefaultRowHeightInPoints(24);

					CSVFormat format = CSVFormat.DEFAULT.withIgnoreEmptyLines(false);

					// 生成并设置另一个样式 内容的背景
					CellStyle style = workbook.createCellStyle();
					// style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
					// style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					// style.setBorderBottom(BorderStyle.THICK);
					// style.setBorderLeft(BorderStyle.THICK);
					// style.setBorderRight(BorderStyle.THICK);
					// style.setBorderTop(BorderStyle.THICK);
					// style.setAlignment(HorizontalAlignment.CENTER);
					// style.setVerticalAlignment(VerticalAlignment.CENTER);
					// 生成另一个字体
					Font font = workbook.createFont();
					// font.setFontName("Arial");
					font.setFontName("等线");
					font.setFontHeight((short) 220);
					font.setBold(false);
					// 把字体应用到当前的样式
					style.setFont(font);

					CellStyle errorStyle = workbook.createCellStyle();
					errorStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
					errorStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
					errorStyle.setFillPattern(FillPatternType.ALT_BARS);
					errorStyle.setFont(font);

					CSVParser parser = CSVParser.parse(f.toFile(), Charset.defaultCharset(), format);
					Map<String, Institution> ins = namingService.get(Database.InCites);
					int i = 0;
					for (CSVRecord record : parser.getRecords()) {
						Row row = sheet.createRow(i);
						if (record.size() == 0) {

						} else if (record.size() == 1) {
							row.createCell(0).setCellValue(new XSSFRichTextString(record.get(0)));
						} else {
							int j = 0;
							for (String str : record) {

								Cell cell;
								if (j == 0) {
									cell = row.createCell(j);
									if (i == 0) {
										Cell c1 = row.createCell(1);
										c1.setCellStyle(style);
										c1.setCellValue("名称（中文）");

										Cell c2 = row.createCell(2);
										c2.setCellStyle(style);
										c2.setCellValue("数字代码");
									} else {
										Institution in = ins.get(str);
										if (in == null) {
											Cell c1 = row.createCell(1);
											c1.setCellStyle(errorStyle);
											Cell c2 = row.createCell(2);
											c2.setCellStyle(errorStyle);
										} else {
											Cell c1 = row.createCell(1);
											c1.setCellStyle(style);
											c1.setCellValue(in.getName());

											Cell c2 = row.createCell(2);
											c2.setCellStyle(style);
											if (in.getCode() != null) {
												c2.setCellValue(NumberUtils.toDouble(in.getCode()));
											}
										}
									}
								} else {
									cell = row.createCell(j + 2);
								}
								cell.setCellStyle(style);
								if (NumberUtils.isCreatable(str)) {
									cell.setCellValue(NumberUtils.toDouble(str));
								} else {
									RichTextString text = new XSSFRichTextString(str);
									cell.setCellValue(text);
								}
								j++;
							}
						}
						// System.out.println();
						i++;
					}
					workbook.write(out);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			}
		});
	}

	public static void main(String[] args) throws IOException {
		new ReadCsvToExcel().readCsvAndTurn();
	}
}
