package com.rubyren.usnews;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.commons.lang3.math.NumberUtils;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.HttpClients;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

import lombok.val;

public class UsnewsCraw1 {
	public static void main(String[] args) throws Exception {

		ObjectMapper om = new ObjectMapper();
		Path p = Paths.get("/Users/zhouning/git/excel_combine/src/main/resources/usnews/subject.txt");
		val to = Paths.get("/Users/zhouning/workspace/eclipse-rubyren-1/patent/target/usnews.xlsx");

		val httpclient = HttpClients.createDefault();

		val subjects = Files.readAllLines(p);
		try (val workbook = new SXSSFWorkbook(); val out = Files.newOutputStream(to)) {
			A: for (val subject : subjects) {
				val sheet = workbook.createSheet(subject);
				sheet.setDefaultColumnWidth(12);
				val head = sheet.createRow(0);
				head.createCell(0).setCellValue("学校");
				head.createCell(1).setCellValue("国家");
				head.createCell(2).setCellValue("省");
				head.createCell(3).setCellValue("市");
				head.createCell(4).setCellValue("分数");
				head.createCell(5).setCellValue("全局排名");
				head.createCell(6).setCellValue("学科排名");
				int lineNum = 1;

				B: for (int i = 1; i <= 125; i++) {
					HttpGet httpget = new HttpGet("https://www.usnews.com/education/best-global-universities/"
							+ subject + "?page=" + i  + "&format=json");
					httpget.addHeader("User-Agent",
							"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.90 Safari/537.36");
					httpget.addHeader("referer",
							"https://www.usnews.com/education/best-global-universities/rankings?page=1&format=json");
					httpget.addHeader("cookie",
							"ak_bmsc=19175611F5D00386BA117EB987D7F5BF1703680B2C29000095D4095D0B09E529~plUsrk8hVBPUv4Dx08Z8tkOQ047E+wxGqpel5Nnbb9N0VsVs2VMwHBM3xsgNj9Q9TJ/xkvnfuqxp+DIkHqX1zPmArWOYHkAEt8JXh8jUNKRuxXL9OZDFvg+GsB8VX9y5ltpcTL9mqCmm6DrpHCm91ZFVmnUEEpiY6Gu4WgyJbkh9nYYy/9lqfBy7JaqHQ6VfU7w6YuTprBZSXRfr/bG1HqINyQ6yGdx7WmLB47kHHtKS2F68lOnkBghmKro34jat+y; bm_sv=A0C35C538CE25106090B06396A897F49~FK2Mwn17pw6a7JLoY3j7wqrXugx9nyIexMXD+82/m1ONhW1tMUEaW7IPVcdxT866YYTcDJ4E+/pT9bN7YuK90+mM0skK96QgXKPS+D0eStgUI1vbzGIKhSziZX+/5M/zhfL2TEAJflCxMqJDeJlaQymQuvK0/5yWmYvcCLB1ArI=; akacd_www=2177452799~rv=57~id=31d5590f77de3bd08dcaefe4b6545298");
					val response = httpclient.execute(httpget);
					int statue = response.getStatusLine().getStatusCode();
					if (statue < 200 || statue >= 300) {
						System.out.println(statue);
						continue;
					}
					val node = om.readTree(response.getEntity().getContent());
					val result = node.get("results");
					for (val r : result) {
						val row = sheet.createRow(lineNum);
						row.createCell(0).setCellValue(r.get("name").asText());
						row.createCell(1).setCellValue(r.get("country_name").asText());
						row.createCell(2).setCellValue(r.get("country_subdivision").asText());
						row.createCell(3).setCellValue(r.get("city").asText());
						row.createCell(4).setCellValue(r.get("score").asText());
						row.createCell(5).setCellValue(r.get("global_rank").asText());
						row.createCell(6).setCellValue(r.get("rank").asText());
						lineNum++;
					}
					String pages = node.get("pagination").get("last_page").asText();
					if (NumberUtils.toInt(pages) <= i) {
						break B;
					}
				}

			}
			workbook.write(out);
		}
	}
}
