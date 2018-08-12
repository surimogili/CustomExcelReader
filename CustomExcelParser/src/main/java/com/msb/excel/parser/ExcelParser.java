package com.msb.excel.parser;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Sheet;

import com.msb.excel.parser.annotation.ExcelCellName;
import com.msb.excel.parser.annotation.ExcelObject;
import com.msb.excel.parser.annotation.MappedExcelObject;
import com.msb.excel.parser.annotation.ParseType;

public class ExcelParser {

	private static final Log LOG = LogFactory.getLog(ExcelParser.class);
	private Map<String, Integer> titles;
	private int HEADER_ROW = 0;
	private Sheet sheet;

	public ExcelParser(Sheet sheet) {
		this.sheet = sheet;
		this.titles = getTitles(sheet, HEADER_ROW);
	}

	public <T> List<T> createEntity(Class<T> clazz) {
		List<T> list = new ArrayList<>();
		ExcelObject excelObject = getExcelObject(clazz);

		int end = getEnd(clazz, excelObject);

		for (int currentLocation = excelObject.start(); currentLocation <= end; currentLocation++) {
			T object = getNewInstance(clazz, excelObject.parseType(), currentLocation, false, null);
			List<Field> mappedExcelFields = getMappedExcelObjects(clazz);
			for (Field mappedField : mappedExcelFields) {
				Class<?> fieldType = mappedField.getType();
                Class<?> clazz1 = fieldType.equals(List.class) ? getFieldType(mappedField) : fieldType;
                List<T> filedValues = new ArrayList<>();
                ExcelObject excelObjectMap = getExcelObject(clazz1);
                boolean loop = excelObjectMap.loop();
                int loopLength = excelObjectMap.looplength();
                for(int i=1; i<=loopLength; i++)
                {
                	T objectMap = (T) getNewInstance(clazz1, excelObject.parseType(), currentLocation, loop, Integer.toString(i));
                	filedValues.add(objectMap);
                }
                setFieldValue(mappedField, object, filedValues);
			}
			list.add(object);
		}

		return list;
	}

	private <T> T getNewInstance(Class<T> clazz, ParseType parseType, Integer currentLocation, boolean loop, String loopLength) {
		T object = getInstance(clazz);
		Map<String, Field> excelFieldNamePositionMap = getExcelNameFieldPositionMap(clazz);
		for (String excelFieldName : excelFieldNamePositionMap.keySet()) {
			Field field = excelFieldNamePositionMap.get(excelFieldName);
			if(loop)
			{
				excelFieldName = excelFieldName+loopLength;
			}
			int fieldPosition = titles.get(excelFieldName);
			Object cellValue = null;
			if (ParseType.ROW == parseType) {
				cellValue = HSSFHelper.getCellValue(sheet, field.getType(), currentLocation, fieldPosition, false);
			}
			setFieldValue(field, object, cellValue);
		}
		return object;
	}

	private <T> List<Field> getMappedExcelObjects(Class<T> clazz) {
		List<Field> fieldList = new ArrayList<>();
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			MappedExcelObject mappedExcelObject = field.getAnnotation(MappedExcelObject.class);
			if (mappedExcelObject != null) {
				field.setAccessible(true);
				fieldList.add(field);
			}
		}
		return fieldList;
	}

	private <T> void setFieldValue(Field field, T object, Object cellValue) {
		try {
			field.set(object, cellValue);
		} catch (IllegalArgumentException | IllegalAccessException e) {
			LOG.error("Exception occurred while setting field value ", e);
		}
	}

	private <T> Map<String, Field> getExcelNameFieldPositionMap(Class<T> clazz) {
		Map<String, Field> fieldMap = new HashMap<>();
		return fillMap(clazz, fieldMap);
	}

	private <T> Map<String, Field> fillMap(Class<T> clazz, Map<String, Field> fieldMap) {
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			ExcelCellName excelCellName = field.getAnnotation(ExcelCellName.class);
			if (excelCellName != null) {
				field.setAccessible(true);
				fieldMap.put(excelCellName.value(), field);
			}
		}
		return fieldMap;
	}

	private <T> T getInstance(Class<T> clazz) {
		T object;
		try {
			Constructor<T> constructor = clazz.getDeclaredConstructor();
			constructor.setAccessible(true);
			object = constructor.newInstance();
		} catch (Exception e) {
			LOG.error("Exception occurred while instantiating the class " + clazz.getName(), e);
			return null;
		}
		return object;
	}

	private Map<String, Integer> getTitles(Sheet sheet, int rowNum) {
		Map<String, Integer> titlesMap = new HashMap<>();
		int colNum = sheet.getRow(rowNum).getLastCellNum();
		if (sheet.getRow(rowNum).cellIterator().hasNext()) {
			for (int i = 0; i < colNum; i++) {
				titlesMap.put(sheet.getRow(rowNum).getCell(i).getStringCellValue(), i);
			}
		}
		return titlesMap;
	}

	private <T> ExcelObject getExcelObject(Class<T> clazz) {
		ExcelObject excelObject = clazz.getAnnotation(ExcelObject.class);
		if (excelObject == null) {
			LOG.error("Invalid class configuration - ExcelObject annotation missing - " + clazz.getSimpleName());
		}
		return excelObject;
	}

	private <T> int getEnd(Class<T> clazz, ExcelObject excelObject) {
		int end = excelObject.end();
		if (end > 0) {
			return end;
		}
		return getRowOrColumnEnd(sheet, clazz);
	}

	public <T> int getRowOrColumnEnd(Sheet sheet, Class<T> clazz) {
		int maxCellNumber = 0;
		ExcelObject excelObject = getExcelObject(clazz);
		ParseType parseType = excelObject.parseType();
		if (parseType == ParseType.ROW) {
			maxCellNumber = sheet.getLastRowNum() + 1;
		}
		return maxCellNumber;
	}
	
	 private Class<?> getFieldType(Field field) {
	        Type type = field.getGenericType();
	        if (type instanceof ParameterizedType) {
	            ParameterizedType pt = (ParameterizedType) type;
	            return (Class<?>) pt.getActualTypeArguments()[0];
	        }

	        return null;
	    }
}
