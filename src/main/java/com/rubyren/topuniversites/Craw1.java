package com.rubyren.topuniversites;

import java.io.OutputStream;
import java.io.Serializable;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.Data;
import lombok.val;

public class Craw1 {
	public static void main(String[] args) throws Exception {
		val om = new ObjectMapper();
//		Path p2 = Paths.get("/Users/zhouning/git/excel_combine/src/main/resources/topuniversites/2.json");
//		val warp2 = om.readValue(Files.newInputStream(p2), DataWarp2.class);
//		warp2.getColumns().forEach(c -> {
//			c.setName(Jsoup.parse(c.getTitle()).text());
//		});
//		warp2.getColumns().forEach(c -> {
//			if (c.getData().endsWith("_rank")) {
//				//TODO
//			}
//			if (c.getData().endsWith("_rank_d")) {
//				//TODO
//			}
//		});
//
//		val buff = new HashMap<String, HashMap<String, String>>();
//		for (val data : warp2.getData()) {
//			val uni = new HashMap<String, String>();
//			for (val c : warp2.getColumns()) {
//				uni.put(c.getName(), Jsoup.parse(data.get(c.getData()).asText()).text());
//			}
//			val uniHtml = Jsoup.parse(data.get("uni").asText());
//			val uniUrl = uniHtml.select("a[href]").attr("href");
//			buff.put(uniUrl, uni);
//
//		}

		Path p1 = Paths.get("/Users/zhouning/git/excel_combine/src/main/resources/topuniversites/1.json");
		val warp1 = om.readValue(Files.newInputStream(p1), DataWarp1.class);
		val data = warp1.getData();
		val to = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/patent/target/to.xlsx");
		Files.deleteIfExists(to);
		Files.createDirectories(to.getParent());
		try (Workbook workbook = new SXSSFWorkbook(); OutputStream out = Files.newOutputStream(to)) {

			Sheet sheet = workbook.createSheet();
			sheet.setDefaultColumnWidth(9);
			sheet.setDefaultRowHeightInPoints(24);
			Row head = sheet.createRow(0);
			head.createCell(0).setCellValue("名称");
			head.createCell(1).setCellValue("区域");
			head.createCell(2).setCellValue("国家");
			head.createCell(3).setCellValue("评分");
			head.createCell(4).setCellValue("排名");
			head.createCell(5).setCellValue("星级");
			int i = 1;
			for (val school : data) {
				Row row = sheet.createRow(i);
				row.createCell(0).setCellValue(school.getTitle());
				row.createCell(1).setCellValue(school.getRegion());
				row.createCell(2).setCellValue(school.getCountry());
				row.createCell(3).setCellValue(school.getScore());
				row.createCell(4).setCellValue(school.getRankDisplay());
				row.createCell(5).setCellValue(school.getStars());
//				String rank_display = school.get("rank_display").asText();
//				String title = school.get("title").asText();
//				String score = school.get("score").asText();
//				String country = school.get("country").asText();
//				String region = school.get("region").asText();
//				String stars = school.get("stars").asText();

//				System.out.println(rank_display + "\t:\t" + title + "\t:\t" + score + "\t:\t" + country + "\t:\t"
//						+ region + "\t:\t" + stars);

				i++;
			}
			workbook.write(out);
		}
		System.out.println();
	}

	@Data
	@JsonInclude(Include.ALWAYS)
	private static class DataWarp1 implements Serializable {
		private static final long serialVersionUID = -2051375493445184925L;
		private List<Data1> data;
	}

	@Data
	@JsonInclude(Include.ALWAYS)
	private static class Data1 implements Serializable {
		private static final long serialVersionUID = -6936072830782908763L;
		private String nid;
		private String url;
		private String title;
		private String logo;
		@JsonProperty("core_id")
		private String coreId;
		private String score;
		@JsonProperty("rank_display")
		private String rankDisplay;
		private String country;
		private String region;
		private String stars;
		private String guide;
	}

	@Data
	@JsonInclude(Include.ALWAYS)
	private static class DataWarp2 implements Serializable {
		private static final long serialVersionUID = -7562483375452442151L;
		private List<Column1> columns;
		private List<ObjectNode> data;
	}

	@Data
	@JsonInclude(Include.ALWAYS)
	private static class Column1 implements Serializable {
		private static final long serialVersionUID = -494488129963477549L;
		private String data;
		private String title;
		private String visible;
		private String className;
		private String searchable;
		private String type;
		private String orderSequence;
		private String orderable;

		@JsonIgnore
		private String name;
	}
}
