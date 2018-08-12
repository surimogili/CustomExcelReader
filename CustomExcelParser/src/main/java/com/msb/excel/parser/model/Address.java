package com.msb.excel.parser.model;

import com.msb.excel.parser.annotation.ExcelCellName;
import com.msb.excel.parser.annotation.ExcelObject;
import com.msb.excel.parser.annotation.ParseType;

@ExcelObject(parseType = ParseType.ROW, start=1, loop=true, looplength=2)
public class Address {

	@ExcelCellName("City")
	private String city;
	
	@ExcelCellName("State")
	private String state;

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}
	
	
}
