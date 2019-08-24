package com.muscleape.excel.support;

import lombok.Getter;

/**
 * 标识字段类型
 *
 * @author Muscleape
 * @date 2019/8/24 11:29
 * @return
 */

@Getter
public enum FieldDataTypeEnum {

    /**
     * 标识为普通类型，也兼容改版之前数据
     */
    NORMAL("NORMAL"),
    /**
     * 标识为日期类型，需要根据后面指定的格式格式化日期数据
     */
    DATE("DATE"),
    /**
     * 标识为键值对类型，需要根据指定的JSON格式键值对转换数据
     */
    MAP("MAP");

    private String dataTypeName;

    FieldDataTypeEnum(String dataTypeName) {
        this.dataTypeName = dataTypeName;
    }
}
