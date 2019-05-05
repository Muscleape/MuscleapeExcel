package com.muscleape.excel.annotation;

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
    String format() default "";

    /**
     * according the JSON convert key to value;
     * =====================================
     * Default JSON format: {'k1':'v1','k2':'v2'}
     *
     * @return
     * @author Muscleape
     */
    String keyValue() default "";
}
