package com.muscleape.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

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
