package com.msb.excel.parser;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.msb.excel.parser.model.Address;
import com.msb.excel.parser.model.Person;

/**
 * Hello world!
 *
 */
public class App 
{
	static InputStream inputStream;
    public static void main( String[] args ) throws Exception
    {
    	String SAMPLE_PERSON_DATA_FILE_PATH = "src/main/resources/person.xlsx"; 
    	File file = new File(SAMPLE_PERSON_DATA_FILE_PATH);
        InputStream inputStream = new FileInputStream(file);
        ExcelParser parser = new ExcelParser(openSheet(inputStream, file));
        
        List<Person> list = parser.createEntity(Person.class);
        for(Person person : list)
        {
        	String msg = "Name : "+ person.getName();
        	String addreMsg = null;
        	for(Address addre : person.getAddresses())
        	{
        		if(addreMsg != null)
        			addreMsg = addreMsg + " ,City: "+ addre.getCity() + ", State: "+ addre.getState();
        		else
        			addreMsg =  "City: "+ addre.getCity() + ", State: "+ addre.getState();
        			
        	}
        	System.out.println(msg + ", " + addreMsg);
        }
    }
    
    public static Sheet openSheet(InputStream inputStream, File file) throws IOException {
        Workbook workbook;
        if(file.getName().endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook.getSheetAt(0);
    }
}
