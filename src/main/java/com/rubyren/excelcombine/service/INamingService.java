package com.rubyren.excelcombine.service;

import java.util.Map;

import com.rubyren.excelcombine.model.Database;
import com.rubyren.excelcombine.model.Institution;

public interface INamingService {
	public Map<String, Institution> get(Database database);
}
