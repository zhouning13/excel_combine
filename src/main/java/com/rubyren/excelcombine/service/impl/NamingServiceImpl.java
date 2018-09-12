package com.rubyren.excelcombine.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.rubyren.excelcombine.model.Database;
import com.rubyren.excelcombine.model.Institution;
import com.rubyren.excelcombine.service.INamingService;

public class NamingServiceImpl implements INamingService {
	private Map<Database, Map<String, Institution>> map = null;
	private File file = Paths.get("C:\\Users\\周宁\\Desktop\\新建文件夹1\\InCites 中国大陆机构(914所)-迟诚-by 王燕rev - 任国华rev.xlsx")
			.toFile();

	public Map<String, Institution> get(Database database) {
		if (map == null) {
			try (FileInputStream stream = new FileInputStream(file); XSSFWorkbook workbook = new XSSFWorkbook(stream)) {
				XSSFSheet sheet = workbook.getSheetAt(0);
				Map<Database, Map<String, Institution>> m = new HashMap<>();
				for (Row row : sheet) {
					String incites = get(row, 1);
					String esi = get(row, 2);
					String name = get(row, 3);
					String code = get(row, 6);
					Institution institution = new Institution(name, code);

					Map<String, Institution> inCitesMap = m.get(Database.InCites);
					if (inCitesMap == null) {
						inCitesMap = new HashMap<>();
						m.put(Database.InCites, inCitesMap);
					}
					inCitesMap.put(incites, institution);

					Map<String, Institution> esiMap = m.get(Database.ESI);
					if (esiMap == null) {
						esiMap = new HashMap<>();
						m.put(Database.ESI, esiMap);
					}
					esiMap.put(esi, institution);

				}

				map = m;
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return map.get(database);
	}

	private String get(Row row, int i) {
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
		new NamingServiceImpl().get(Database.InCites);
		System.out.println();
	}
}
