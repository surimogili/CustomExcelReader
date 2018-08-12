package com.msb.excel.parser;

import static java.text.MessageFormat.format;

import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.Iterator;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class HSSFHelper {
	
	private static DataFormatter formatter = new DataFormatter();
	private static final Log LOG = LogFactory.getLog(HSSFHelper.class);

    @SuppressWarnings("unchecked")
    public static <T> T getCellValue(Sheet sheet, Class<T> type, Integer row, Integer col, boolean zeroIfNull) {
        Cell cell = getCell(sheet, row, col);

        return validateAndParseValue(cell, sheet.getSheetName(), type, row, col, zeroIfNull);
    }


    public static <T> T getCellValue(Row row, String sheetName, Class<T> type, Integer rowIndex, Integer col, boolean zeroIfNull) {
        Cell cell = row.getCell(col - 1);

        return validateAndParseValue(cell, sheetName, type, rowIndex, col, zeroIfNull);
    }

    @SuppressWarnings("unchecked")
    private static <T> T validateAndParseValue(Cell cell, String sheetName, Class<T> type, Integer row, Integer col, boolean zeroIfNull) {
        if (type.equals(String.class)) {
            return (T) getStringCell(cell);
        }

        if (type.equals(Date.class)) {
            return cell == null ? null : (T) getDateCell(cell, new Locator(sheetName, row, col));
        }

        if (type.equals(LocalDate.class)) {
            return cell == null ? null : (T) getLocalDateCell(cell, new Locator(sheetName, row, col));
        }

        if (type.equals(LocalDateTime.class)) {
            return cell == null ? null : (T) getLocalDateTimeCell(cell, new Locator(sheetName, row, col));
        }

        if (type.equals(Integer.class)) {
            return (T) getIntegerCell(cell, zeroIfNull, new Locator(sheetName, row, col));
        }

        if (type.equals(Double.class)) {
            return (T) getDoubleCell(cell, zeroIfNull, new Locator(sheetName, row, col));
        }

        if (type.equals(Long.class)) {
            return (T) getLongCell(cell, zeroIfNull, new Locator(sheetName, row, col));
        }

        if (type.equals(BigDecimal.class)) {
            return (T) getBigDecimalCell(cell, zeroIfNull, new Locator(sheetName, row, col));
        }

        
        return null;
    }

    private static LocalDate getLocalDateCell(Cell cell, Locator locator) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                LOG.error(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
            }

            Instant instant = Instant.ofEpochMilli(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).getTime());
            LocalDateTime res = LocalDateTime.ofInstant(instant, ZoneId.systemDefault());
            return res.toLocalDate();

        } catch (IllegalStateException illegalStateException) {
            LOG.error(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
        }
        return null;
    }

    private static LocalDateTime getLocalDateTimeCell(Cell cell, Locator locator) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                LOG.error(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
            }

            Instant instant = Instant.ofEpochMilli(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).getTime());
            return LocalDateTime.ofInstant(instant, ZoneId.systemDefault());

        } catch (IllegalStateException illegalStateException) {
            LOG.error(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
        }
        return null;
    }

    private static BigDecimal getBigDecimalCell(Cell cell, boolean zeroIfNull, Locator locator) {
        String val = getStringCell(cell);
        if(val == null || val.trim().equals("")) {
            if(zeroIfNull) {
                return BigDecimal.ZERO;
            }
            return null;
        }
        try {
            return new BigDecimal(val);
        } catch (NumberFormatException e) {
            LOG.error(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
        }

        if (zeroIfNull) {
            return BigDecimal.ZERO;
        }
        return null;
    }

    static Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
        Row row = sheet.getRow(rowNumber);
        return row == null ? null : row.getCell(columnNumber);
    }

    public static Row getRow(Iterator<Row> iterator, int rowNumber) {
        Row row;
        while (iterator.hasNext()) {
            row = iterator.next();
            if (row.getRowNum() == rowNumber - 1) {
                return row;
            }
        }
        throw new RuntimeException("No Row with index: " + rowNumber + " was found");
    }

    static String getStringCell(Cell cell) {
        if (cell == null) {
            return null;
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            int type = cell.getCachedFormulaResultType();

            if (type == HSSFCell.CELL_TYPE_NUMERIC) {
            	FormulaEvaluator fe = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator(); 
            	return formatter.formatCellValue(cell, fe);
            }

            if (type == HSSFCell.CELL_TYPE_ERROR) {
                return "";
            }

            if (type == HSSFCell.CELL_TYPE_STRING) {
                return cell.getRichStringCellValue().getString().trim();
            }

            if (type == HSSFCell.CELL_TYPE_BOOLEAN) {
                return "" + cell.getBooleanCellValue();
            }

        } else if (cell.getCellType() != HSSFCell.CELL_TYPE_NUMERIC) {
            return cell.getRichStringCellValue().getString().trim();
        }

        return formatter.formatCellValue(cell);
    }

    static Date getDateCell(Cell cell, Locator locator) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                LOG.error(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
            }
            return HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
        } catch (IllegalStateException illegalStateException) {
            LOG.error(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
        }
        return null;
    }

    static Double getDoubleCell(Cell cell, boolean zeroIfNull, Locator locator) {
        if (cell == null) {
            return zeroIfNull ? 0d : null;
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC || cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            return cell.getNumericCellValue();
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            return zeroIfNull ? 0d : null;
        }

        LOG.error(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
        return null;
    }

    static Long getLongCell(Cell cell, boolean zeroIfNull, Locator locator) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator);
        return doubleValue == null ? null : doubleValue.longValue();
    }

    static Integer getIntegerCell(Cell cell, boolean zeroIfNull, Locator locator) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator);
        return doubleValue == null ? null : doubleValue.intValue();
    }

    private static Double getNumberWithoutDecimals(Cell cell, boolean zeroIfNull, Locator locator) {
        Double doubleValue = getDoubleCell(cell, zeroIfNull, locator);
        if (doubleValue != null && doubleValue % 1 != 0) {
            LOG.error(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol()));
        }
        return doubleValue;
    }

}
