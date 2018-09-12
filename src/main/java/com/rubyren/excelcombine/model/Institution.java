package com.rubyren.excelcombine.model;

import java.io.Serializable;

public class Institution implements Serializable {
	private static final long serialVersionUID = 8044042139444410395L;
	private String name;
	private String code;

	public Institution() {
		super();
	}

	public Institution(String name, String code) {
		super();
		this.name = name;
		this.code = code;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}

}
