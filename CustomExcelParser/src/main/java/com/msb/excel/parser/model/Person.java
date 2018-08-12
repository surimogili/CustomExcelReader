package com.msb.excel.parser.model;

import java.util.List;

import com.msb.excel.parser.annotation.ExcelCellName;
import com.msb.excel.parser.annotation.ExcelObject;
import com.msb.excel.parser.annotation.MappedExcelObject;
import com.msb.excel.parser.annotation.ParseType;

@ExcelObject(parseType = ParseType.ROW, start=1)
public class Person {

    @ExcelCellName("Name")
    protected String name;

    @ExcelCellName("Address")
    protected String address;

    @ExcelCellName("Mobile")
    protected String mobile;

    @ExcelCellName("Email")
    protected String email;
    
    @MappedExcelObject
    protected List<Address> addresses;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getMobile() {
        return mobile;
    }

    public void setMobile(String mobile) {
        this.mobile = mobile;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

	public List<Address> getAddresses() {
		return addresses;
	}

	public void setAddresses(List<Address> addresses) {
		this.addresses = addresses;
	}
    
}
