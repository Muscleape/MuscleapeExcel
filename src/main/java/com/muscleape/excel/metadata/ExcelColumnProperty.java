package com.muscleape.excel.metadata;

import lombok.Data;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Muscleape
 */
@Data
public class ExcelColumnProperty implements Comparable<ExcelColumnProperty> {

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
    private List<String> head = new ArrayList<String>();

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

    @Override
    public int compareTo(@NotNull ExcelColumnProperty o) {
        int x = this.index;
        int y = o.getIndex();
        return (x < y) ? -1 : ((x == y) ? 0 : 1);
    }
}