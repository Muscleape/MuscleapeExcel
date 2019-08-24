package com.muscleape.excel.annotation;

import com.muscleape.excel.support.FieldDataTypeEnum;

import java.lang.annotation.*;

/**
 * @author Muscleape
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

    /**
     * @return
     */
    String[] value() default {""};


    /**
     * @return
     */
    int index() default 99999;

    /**
     * default @see TypeUtil
     * if default is not  meet you can set format
     *
     * @return
     */
    @Deprecated
    String format() default "";

    /**
     * according the JSON convert key to value;
     * =====================================
     * Default JSON format: {'k1':'v1','k2':'v2'}
     *
     * @return
     * @author Muscleape
     */
    @Deprecated
    String keyValue() default "";

    /**
     * 标识字段类型，已确定字段需要进行哪些处理
     *
     * @return com.muscleape.excel.support.FieldDataTypeEnum
     * @author Muscleape
     * @date 2019/8/24 11:38
     */
    FieldDataTypeEnum fieldDataType() default FieldDataTypeEnum.NORMAL;

    /**
     * 日期数据格式化样式
     * <p>
     * FieldDataTypeEnum.DATE类型的数据，格式化规则
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:53
     */
    String dateFormat() default "";

    /**
     * 标识字段类型
     * FieldDataTypeEnum.MAP类型的数据，键值对转换方式
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:54
     */
    String dateMapJsonStr() default "";

    /**
     * 键值对转换时，key为null时，对应转换值
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:56
     */
    String dateMapKeyNull() default "";

    /**
     * 键值对转换时，key为不存在时
     *
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/8/24 11:56
     */
    String dateMapKeyOther() default "";
}
