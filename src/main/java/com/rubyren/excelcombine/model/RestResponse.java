package com.rubyren.excelcombine.model;

import java.io.Serializable;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonInclude.Include;

@JsonInclude(Include.NON_NULL)
public class RestResponse<T> implements Serializable {
	private static final long serialVersionUID = 8157181600871078769L;
	private int code = 200;
	private String message = "调用成功";
	private T result;

	public RestResponse() {
		super();
	}

	public RestResponse(int code, String message) {
		super();
		this.code = code;
		this.message = message;
	}

	public RestResponse(int code, String message, T result) {
		super();
		this.code = code;
		this.message = message;
		this.result = result;
	}

	public int getCode() {
		return code;
	}

	public void setCode(int code) {
		this.code = code;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

	public T getResult() {
		return result;
	}

	public void setResult(T result) {
		this.result = result;
	}

}
