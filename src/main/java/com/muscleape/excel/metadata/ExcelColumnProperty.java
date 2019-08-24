package com.muscleape.excel.metadata;

import com.muscleape.excel.support.FieldDataTypeEnum;
import lombok.Data;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Muscleape
 */
@Data
// public class ExcelColumnProperty implements Comparable<ExcelColumnProperty> {
public class ExcelColumnProperty {

    /**
     *
     */
    private Field field;

    /**
     *
     */
    private int index = 99999;

    /**
     *
     */
    private List<String> head = new ArrayList<>();

    /**
     * datetime format
     */
    private String format;

    /**
     * according the JSON convert key to value;
     * =====================================
     * Default JSON format: {'k1':'v1','k2':'v2'}
     */
    private String keyValue;

    /**
     * 标识字段类型，已确定字段需要进行哪些处理
     *
     * @return com.muscleape.excel.support.FieldDataTypeEnum
     * @author Muscleape
     * @date 2019/8/24 11:38
     */
    FieldDataTypeEnum fieldDataType;

    /**
     * 日期数据格式化样式
     * <p>
     * FieldDataTypeEnum.DATE类型的数据，格式化规则
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:53
     */
    String dateFormat;

    /**
     * 标识字段类型
     * FieldDataTypeEnum.MAP类型的数据，键值对转换方式
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:54
     */
    String dateMapJsonStr;

    /**
     * 键值对转换时，key为null时，对应转换值
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:56
     */
    String dateMapKeyNull;

    /**
     * 键值对转换时，key为不存在时
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:56
     */
    String dateMapKeyOther;
}