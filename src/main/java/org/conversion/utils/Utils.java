package org.conversion.utils;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * @author Nicholas Curl
 */
public class Utils {

    /**
     * The instance of the logger
     */
    private static final Logger logger = LogManager.getLogger(Utils.class);

    public static int getLastColumn(XSSFSheet sheet) {
        int lastColumn = 0;
        for (Row row : sheet) {
            if ((int) row.getLastCellNum() > lastColumn) {
                lastColumn = row.getLastCellNum();
            }
        }
        return lastColumn;
    }

    public static Object getCellValue(XSSFCell cell, XSSFFormulaEvaluator evaluator) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    }
                    else {
                        return cell.getNumericCellValue();
                    }
                case FORMULA:
                    evaluator.evaluate(cell);
                    switch (cell.getCachedFormulaResultType()) {
                        case STRING:
                            return cell.getStringCellValue().trim();
                        case BOOLEAN:
                            return cell.getBooleanCellValue();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return cell.getDateCellValue();
                            }
                            else {
                                return cell.getNumericCellValue();
                            }
                        default:
                            return null;
                    }
                default:
                    return null;
            }
        }
        else {
            return null;
        }
    }
    public static int findBOMColumn(int itemNum, List<Object> topRow){
        int col = 0;
        for(Object o : topRow){
            if (o instanceof Double) {
                double aDouble = (double) o;
                int aInt = (int) aDouble;
                if(aInt == itemNum){
                    return col;
                }
            }
            else if (o instanceof String) {
                String s = (String) o;
                if(s.equalsIgnoreCase(String.valueOf(itemNum))){
                    return col;
                }
            }
            col++;
        }
        return Integer.MAX_VALUE;
    }

    public static String joinSQLLine(String sql){
        return sql.replaceAll("[\n\r]|\\s{2,}", " ");
    }

    public static void setCellValue(XSSFCell cell, Object object){
        if (object instanceof Integer) {
            int integer = (int) object;
            cell.setCellValue(integer);
        }
        else if (object instanceof Double) {
            double aDouble = (double) object;
            cell.setCellValue(aDouble);
        }
        else if (object instanceof Boolean) {
            boolean aBoolean = (boolean) object;
            cell.setCellValue(aBoolean);
        }
        else if (object instanceof Date) {
            Date date = (Date) object;
            cell.setCellValue(date);
        }
        else if (object instanceof Calendar) {
            Calendar calendar = (Calendar) object;
            cell.setCellValue(calendar);
        }
        else if (object instanceof LocalDate) {
            LocalDate localDate = (LocalDate) object;
            cell.setCellValue(localDate);
        }
        else if (object instanceof LocalDateTime) {
            LocalDateTime localDateTime = (LocalDateTime) object;
            cell.setCellValue(localDateTime);
        }
        else if (object instanceof RichTextString) {
            RichTextString richTextString = (RichTextString) object;
            cell.setCellValue(richTextString);
        }
        else if (object instanceof String) {
            String s = (String) object;
            cell.setCellValue(s);
        }
        else if(object instanceof BigDecimal){
            BigDecimal bigDecimal = (BigDecimal) object;
            cell.setCellValue(bigDecimal.doubleValue());
        }
        else {
            cell.setBlank();
        }
    }
}
