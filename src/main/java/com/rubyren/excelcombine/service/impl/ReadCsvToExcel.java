package com.rubyren.excelcombine.service.impl;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.EncryptedDocumentException;
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
	private SimpleDateFormat df = new SimpleDateFormat("yyyyMMdd");
	private CSVFormat format = CSVFormat.DEFAULT.withIgnoreEmptyLines(false);
	private Map<String, Institution> ins = namingService.get(Database.InCites);

	public void readCsvAndTurn() throws IOException {
		Path base = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201906/");
		Path from = base.resolve("6月趋势数据/");
		Path to = base.resolve("趋势数据to/趋势数据" + df.format(new Date()) + ".xlsx");
		Files.createDirectories(to.getParent());
		Path to2 = base.resolve("趋势数据to/过程数据/");
		Files.createDirectories(to2);
		try (Workbook workbook = new SXSSFWorkbook(); FileOutputStream out = new FileOutputStream(to.toFile())) {

			
			Files.walk(from, 1).sorted().forEach(f -> {
				if (Objects.equals(".DS_Store", f.getFileName().toString())) {
					return;
				}
				if (!Objects.equals(from, f)) {

					try {
						String name = f.getFileName().toString();
						name = name.substring(0, name.length() - 4);
						System.out.println(name);

						Sheet sheet = workbook.createSheet(name);
						sheet.setDefaultColumnWidth(9);
						sheet.setDefaultRowHeightInPoints(24);
						
						deal(workbook,sheet,f);
						
						Workbook progressWorkbook = new SXSSFWorkbook();
						Sheet progressSheet = progressWorkbook.createSheet(name);
						progressSheet.setDefaultColumnWidth(9);
						progressSheet.setDefaultRowHeightInPoints(24);
						deal(progressWorkbook,progressSheet,f);
						FileOutputStream progressOut = new FileOutputStream(to2.resolve(name+".xlsx").toFile());
						progressWorkbook.write(progressOut);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}

				}
			});
			workbook.write(out);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (EncryptedDocumentException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	}

	private void deal(Workbook workbook, Sheet sheet, Path f) throws IOException {

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
	}

	public static void main(String[] args) throws IOException {
		new ReadCsvToExcel().readCsvAndTurn();
	}
}
