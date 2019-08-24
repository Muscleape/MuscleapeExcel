package com.muscleape.excel.test.model;

import com.muscleape.excel.annotation.ExcelProperty;
import com.muscleape.excel.metadata.BaseRowModel;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

/**
 * @Author: Muscleape
 * @Date: 2019-08-24
 * @Description:
 */

@Data
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class MuscleapeWriteModelOldVersion extends BaseRowModel {

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
    @ExcelProperty(value = "Long类型时间戳", format = "yyyy-MM-dd HH:mm:ss")
    private Long timestampLong;
    /**
     * Date日期类型数据
     */
    @ExcelProperty(value = "Date日期类型", format = "yyyy-MM-dd HH:mm:ss")
    private Date date;
    /**
     * 字符串日期类型
     */
    @ExcelProperty(value = "字符串日期")
    private String dateStr;
    /**
     * 键值对转换key为数字
     */
    @ExcelProperty(value = "键值对转换key为数字", keyValue = "{1:'第一个值',2:'第二个值',3:'第三个值'}")
    private Integer intKeyValue;
    /**
     * 键值对转换key为数字,但是转换JSON中为字符
     */
    @ExcelProperty(value = "键值对转换key为数字,但是转换JSON中为字符", keyValue = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private Integer intStrKeyValue;
    /**
     * 键值对转换key为字符串
     */
    @ExcelProperty(value = "键值对转换key为字符串", keyValue = "{'1':'第一个值','2':'第二个值','3':'第三个值'}")
    private String strKeyValue;
    /**
     * 键值对转换空字符串
     */
    @ExcelProperty(value = "键值对转换空字符串", keyValue = "")
    private String emptyKeyValue;
}
