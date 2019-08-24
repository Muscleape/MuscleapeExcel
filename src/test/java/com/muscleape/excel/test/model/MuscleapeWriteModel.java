package com.muscleape.excel.test.model;

import com.muscleape.excel.annotation.ExcelProperty;
import com.muscleape.excel.metadata.BaseRowModel;
import com.muscleape.excel.support.FieldDataTypeEnum;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * @Author: Muscleape
 * @Date: 2019-08-23
 * @Description:
 */

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class MuscleapeWriteModel extends BaseRowModel {

    /**
     * 整型数据
     */
    @ExcelProperty(value = "主键")
    private Integer id;
    /**
     * 字符串类型数据
     */
    @ExcelProperty(value = "普通字符串类型")
    private String str;
    /**
     * Long类型时间戳
     */
    @ExcelProperty(value = "Long类型时间戳", fieldDataType = FieldDataTypeEnum.DATE, dateFormat = "yyyyMMdd")
    private Long timestampLong;
    /**
     * String类型时间戳
     */
    @ExcelProperty(value = "String类型时间戳", fieldDataType = FieldDataTypeEnum.DATE, dateFormat = "yyyyMMdd")
    private String timestampString;
    /**
     * Date日期类型数据
     */
    @ExcelProperty(value = "Date日期类型", fieldDataType = FieldDataTypeEnum.DATE, dateFormat = "yyyyMMdd")
    private Date date;
    /**
     * 字符串日期类型
     */
    @ExcelProperty(value = "字符串日期", fieldDataType = FieldDataTypeEnum.DATE, dateFormat = "yyyyMMdd")
    private String dateStr;
    /**
     * 键值对转换key为数字
     */
    @ExcelProperty(value = "键值对转换key为数字", fieldDataType = FieldDataTypeEnum.MAP, dateMapJsonStr = "{1:'第一个值',2:'第二个值',3:'第三个值'}")
    private Integer intKeyValue;
    /**
     * 键值对转换key为数字,但是转换JSON中为字符
     */
    @ExcelProperty(value = "键值对转换key为数字,但是转换JSON中为字符", fieldDataType = FieldDataTypeEnum.MAP, dateMapJsonStr = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private Integer intStrKeyValue;
    /**
     * 键值对转换key为字符串
     */
    @ExcelProperty(value = "键值对转换key为字符串", fieldDataType = FieldDataTypeEnum.MAP, dateMapJsonStr = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private String strKeyValue;
    /**
     * 键值对转换,JSON空字符串
     */
    @ExcelProperty(value = "键值对转换,JSON空字符串", fieldDataType = FieldDataTypeEnum.MAP, dateMapJsonStr = "")
    private String emptyKeyValue;
    /**
     * 键值对转换,不存在的值，不指定默认值
     */
    @ExcelProperty(value = "键值对转换,不存在的值,不指定默认值", fieldDataType = FieldDataTypeEnum.MAP, dateMapJsonStr = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private String otherKeyValueNoDefault;
    /**
     * 键值对转换,不存在的值,指定默认值
     */
    @ExcelProperty(value = "键值对转换,不存在的值,指定默认值", fieldDataType = FieldDataTypeEnum.MAP, dateMapKeyOther = "不存在的值", dateMapJsonStr = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private String otherKeyValueWithDefault;
    /**
     * 键值对转换,null值,不指定默认值
     */
    @ExcelProperty(value = "键值对转换,JSON空字符串,不指定默认值", fieldDataType = FieldDataTypeEnum.MAP, dateMapJsonStr = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private String nullKeyValueNoDefault;
    /**
     * 键值对转换,null值,指定默认值
     */
    @ExcelProperty(value = "键值对转换,JSON空字符串,指定默认值", fieldDataType = FieldDataTypeEnum.MAP, dateMapKeyNull = "空数据", dateMapJsonStr = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private String nullKeyValueWithDefault;
}
