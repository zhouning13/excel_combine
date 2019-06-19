package com.rubyren.savedrecs;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ReadTxtAndWriteExcel {

	public static void main(String[] args) throws IOException {
		Path p = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/下载补充数据20190403/");

		Workbook workbook = new SXSSFWorkbook();
		Sheet sheet = workbook.createSheet("sheet1");
		sheet.setDefaultColumnWidth(9);
		sheet.setDefaultRowHeightInPoints(24);

		CellStyle style = workbook.createCellStyle();
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

		Path target = Paths
				.get("/Users/zhouning/workspace/eclipse-rubyren-1/excels/2019/201904/下载补充数据20190403.xlsx");
		Files.deleteIfExists(target);
		OutputStream out = Files.newOutputStream(target);
		int i = 0;
		Path[] files= Files.walk(p).sorted().toArray(Path[]::new);
		
		for(Path f:files) {
			if (Files.isDirectory(f)) {
				continue;
			}
			if (".DS_Store".equalsIgnoreCase(f.getFileName().toString())) {
				continue;
			}
			try {
				List<String> lines = Files.readAllLines(f);
				lines.remove(0);
				int index=0;
				for (String line : lines) {
					Row row = sheet.createRow(i);
					String[] fs = line.split("\t");
					for (int j = 0; j < fs.length; j++) {
						Cell cell = row.createCell(j);
						cell.setCellStyle(style);
						try {
							cell.setCellValue(fs[j]);
						} catch (IllegalArgumentException e) {
							System.out.println(f.getFileName().toString() + "\t" + index +"\t"
									+j+ "\tvalue to long");
						}
					}

//					System.out.println(f.getFileName().toString() + "\t" + i);
					i++;
					index++;
				}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		};
		workbook.write(out);
	}
}
