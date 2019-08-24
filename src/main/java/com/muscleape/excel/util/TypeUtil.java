package com.muscleape.excel.util;

import com.alibaba.fastjson.JSONObject;
import com.muscleape.excel.metadata.ExcelColumnProperty;
import com.muscleape.excel.metadata.ExcelHeadProperty;
import com.muscleape.excel.support.FieldDataTypeEnum;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.springframework.util.ObjectUtils;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Muscleape
 */
public class TypeUtil {

    private static List<String> DATE_FORMAT_LIST = new ArrayList<>(4);

    static {
        DATE_FORMAT_LIST.add("yyyy/MM/dd HH:mm:ss");
        DATE_FORMAT_LIST.add("yyyy-MM-dd HH:mm:ss");
        DATE_FORMAT_LIST.add("yyyyMMdd HH:mm:ss");
    }

    private static int getCountOfChar(String value, char c) {
        int n = 0;
        if (value == null) {
            return 0;
        }
        char[] chars = value.toCharArray();
        for (char cc : chars) {
            if (cc == c) {
                n++;
            }
        }
        return n;
    }

    public static Object convert(String value, Field field, String format, boolean us) {
        if (!StringUtils.isEmpty(value)) {
            if (Float.class.equals(field.getType())) {
                return Float.parseFloat(value);
            }
            if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
                return Integer.parseInt(value);
            }
            if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
                if (null != format && !"".equals(format)) {
                    int n = getCountOfChar(value, '0');
                    return Double.parseDouble(TypeUtil.formatFloat0(value, n));
                } else {
                    return Double.parseDouble(TypeUtil.formatFloat(value));
                }
            }
            if (Boolean.class.equals(field.getType()) || boolean.class.equals(field.getType())) {
                String valueLower = value.toLowerCase();
                if (valueLower.equals("true") || valueLower.equals("false")) {
                    return Boolean.parseBoolean(value.toLowerCase());
                }
                Integer integer = Integer.parseInt(value);
                if (integer == 0) {
                    return false;
                } else {
                    return true;
                }
            }
            if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
                return Long.parseLong(value);
            }
            if (Date.class.equals(field.getType())) {
                if (value.contains("-") || value.contains("/") || value.contains(":")) {
                    return getSimpleDateFormatDate(value, format);
                } else {
                    Double d = Double.parseDouble(value);
                    return HSSFDateUtil.getJavaDate(d, us);
                }
            }
            if (BigDecimal.class.equals(field.getType())) {
                return new BigDecimal(value);
            }
            if (String.class.equals(field.getType())) {
                return formatFloat(value);
            }

        }
        return null;
    }

    public static Boolean isNum(Field field) {
        if (field == null || field.equals("")) {
            return false;
        }
        if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            return true;
        }
        if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
            return true;
        }

        if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
            return true;
        }

        if (BigDecimal.class.equals(field.getType())) {
            return true;
        }
        return false;
    }

    public static Boolean isEmptyJsonObject(String keyValue) {
        JSONObject jsonObject = new JSONObject();
        try {
            jsonObject = JSONObject.parseObject(keyValue);
        } catch (Exception e) {
        }
        return (ObjectUtils.isEmpty(jsonObject) || jsonObject.size() == 0) ? true : false;
    }


    public static Boolean isNum(Object cellValue) {
        if (cellValue instanceof Integer
                || cellValue instanceof Double
                || cellValue instanceof Short
                || cellValue instanceof Long
                || cellValue instanceof Float
                || cellValue instanceof BigDecimal) {
            return true;
        }
        return false;
    }

    public static String getDefaultDateString(Date date) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return simpleDateFormat.format(date);

    }

    public static Date getSimpleDateFormatDate(String value, String format) {
        if (!StringUtils.isEmpty(value)) {
            Date date = null;
            if (!StringUtils.isEmpty(format)) {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
                try {
                    date = simpleDateFormat.parse(value);
                    return date;
                } catch (ParseException e) {
                }
            }
            for (String dateFormat : DATE_FORMAT_LIST) {
                try {
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormat);
                    date = simpleDateFormat.parse(value);
                } catch (ParseException e) {
                }
                if (date != null) {
                    break;
                }
            }

            return date;

        }
        return null;

    }


    public static String formatFloat(String value) {
        if (null != value && value.contains(".")) {
            if (isNumeric(value)) {
                try {
                    BigDecimal decimal = new BigDecimal(value);
                    BigDecimal setScale = decimal.setScale(10, BigDecimal.ROUND_HALF_DOWN).stripTrailingZeros();
                    return setScale.toPlainString();
                } catch (Exception e) {
                }
            }
        }
        return value;
    }

    public static String formatFloat0(String value, int n) {
        if (null != value && value.contains(".")) {
            if (isNumeric(value)) {
                try {
                    BigDecimal decimal = new BigDecimal(value);
                    BigDecimal setScale = decimal.setScale(n, BigDecimal.ROUND_HALF_DOWN);
                    return setScale.toPlainString();
                } catch (Exception e) {
                }
            }
        }
        return value;
    }

    public static final Pattern pattern = Pattern.compile("[\\+\\-]?[\\d]+([\\.][\\d]*)?([Ee][+-]?[\\d]+)?$");

    private static boolean isNumeric(String str) {
        Matcher isNum = pattern.matcher(str);
        if (!isNum.matches()) {
            return false;
        }
        return true;
    }

    public static String formatDate(Date cellValue, String format) {
        SimpleDateFormat simpleDateFormat;
        if (StringUtils.isNotBlank(format)) {
            simpleDateFormat = new SimpleDateFormat(format);
        } else {
            simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        }
        return simpleDateFormat.format(cellValue);
    }

    public static String formatDateTimeStamp(Long cellValue, String format) {
        SimpleDateFormat simpleDateFormat;
        if (StringUtils.isNotBlank(format)) {
            simpleDateFormat = new SimpleDateFormat(format);
        } else {
            simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        }
        return simpleDateFormat.format(cellValue);
    }

    public static String getFieldStringValue(BeanMap beanMap, ExcelColumnProperty excelHeadProperty) {

        // 字段名称
        String fieldName = excelHeadProperty.getField().getName();
        // 原日期格式
        String format = excelHeadProperty.getFormat();
        // 原键值对转换
        String keyValue = excelHeadProperty.getKeyValue();
        // 数据类型
        FieldDataTypeEnum fieldDataType = excelHeadProperty.getFieldDataType();
        // 日期格式
        String dateFormat = excelHeadProperty.getDateFormat();
        // 键值对JSON字符串
        String dateMapJsonStr = excelHeadProperty.getDateMapJsonStr();
        // 键值对中键为null的处理
        String dateMapKeyNull = excelHeadProperty.getDateMapKeyNull();
        // 键值对中键为不存在的值时的处理
        String dateMapKeyOther = excelHeadProperty.getDateMapKeyOther();

        Object value = beanMap.get(fieldName);
        String cellValue;
        switch (fieldDataType) {
            case NORMAL:
                cellValue = handleOldVersionData(value, format, keyValue);
                break;
            case MAP:
                cellValue = handleMapData(value, dateMapJsonStr, dateMapKeyNull, dateMapKeyOther);
                break;
            case DATE:
                cellValue = handleDateData(value, dateFormat);
                break;
            default:
                cellValue = value.toString();
                break;
        }
        return cellValue;
    }

    /**
     * 处理日期类型数据
     *
     * @param value
     * @param dateFormat
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 14:57
     */
    private static String handleDateData(Object value, String dateFormat) {
        if (value == null) {
            return value.toString();
        }
        if (value instanceof Date) {
            return TypeUtil.formatDate((Date) value, dateFormat);
        }
        if (value instanceof Long) {
            return TypeUtil.formatDateTimeStamp((Long) value, dateFormat);
        }
        if (value instanceof String && isNumberFormat(value.toString())) {
            // 字符串类型时间戳
            return TypeUtil.formatDateTimeStamp(Long.valueOf(value.toString()), dateFormat);
        }
        if (value instanceof String) {
            DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(StringUtils.isBlank(dateFormat) ? "yyyy-MM-dd HH:mm:ss" : dateFormat);
            try {
                LocalTime localTime = LocalTime.parse(value.toString(), dateTimeFormatter);
                return localTime.toString();
            } catch (Exception e) {
            }
        }
        return value.toString();
    }

    /**
     * 处理键值对类型转换数据
     *
     * @param value
     * @param dateMapJsonStr
     * @param dateMapKeyNull
     * @param dateMapKeyOther
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 14:46
     */
    private static String handleMapData(Object value, String dateMapJsonStr, String dateMapKeyNull, String dateMapKeyOther) {
        if (StringUtils.isNotBlank(dateMapJsonStr)) {
            JSONObject jsonObject = JSONObject.parseObject(dateMapJsonStr);
            if (value == null) {
                return StringUtils.isBlank(dateMapKeyNull) ? "null" : dateMapKeyNull;
            }
            if (jsonObject.containsKey(value) || jsonObject.containsKey(value.toString())) {
                return jsonObject.get(value) == null ? jsonObject.get(value.toString()).toString() : jsonObject.get(value).toString();
            } else {
                return StringUtils.isBlank(dateMapKeyOther) ? value.toString() : dateMapKeyOther;
            }
        }
        return value.toString();
    }

    /**
     * 兼容旧版本数据，不指定数据类型时
     *
     * @param value
     * @param format
     * @param keyValue
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 14:34
     */
    private static String handleOldVersionData(Object value, String format, String keyValue) {
        String cellValue = value.toString();
        if (value != null) {
            try {
                if (StringUtils.isNotBlank(format)) {
                    if (value instanceof Date) {
                        cellValue = TypeUtil.formatDate((Date) value, format);
                    } else if ((value instanceof Long) && StringUtils.isNotBlank(format)) {
                        // 时间戳类型时间格式化
                        cellValue = TypeUtil.formatDateTimeStamp((Long) value, format);
                    }
                }
                if (StringUtils.isNotBlank(keyValue)) {
                    JSONObject jsonObject = JSONObject.parseObject(keyValue);
                    if (jsonObject.containsKey(value) || jsonObject.containsKey(value.toString())) {
                        cellValue = jsonObject.get(value) == null ? jsonObject.get(value.toString()).toString() : jsonObject.get(value).toString();
                    } else {
                        cellValue = "";
                    }
                }
            } catch (Exception e) {
            }
        }
        return cellValue;
    }

    public static Map getFieldValues(List<String> stringList, ExcelHeadProperty excelHeadProperty, Boolean
            use1904WindowDate) {
        Map map = new HashMap();
        for (int i = 0; i < stringList.size(); i++) {
            ExcelColumnProperty columnProperty = excelHeadProperty.getExcelColumnProperty(i);
            if (columnProperty != null) {
                Object value = TypeUtil.convert(stringList.get(i), columnProperty.getField(),
                        columnProperty.getFormat(), use1904WindowDate);
                if (value != null) {
                    map.put(columnProperty.getField().getName(), value);
                }
            }
        }
        return map;
    }

    /**
     * 是否是数字格式
     *
     * @param number
     * @return boolean
     * @author Muscleape
     * @date 2019/8/24 15:37
     */
    private static boolean isNumberFormat(String number) {
        if (StringUtils.isBlank(number)) {
            return false;
        }
        // 负号
        int minusIndex = number.indexOf("-");
        // 小数点
        int docIndex = number.indexOf(".");
        if (minusIndex > 0) {
            return false;
        }
        if (minusIndex == 0) {
            number = number.substring(1);
        }
        if (docIndex < 0) {
            return StringUtils.isNumeric(number);
        } else {
            String num1 = number.substring(0, docIndex);
            String num2 = number.substring(docIndex + 1);
            return StringUtils.isNumeric(num1) && StringUtils.isNumeric(num2);
        }
    }
}
