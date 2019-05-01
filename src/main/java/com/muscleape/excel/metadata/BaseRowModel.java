package com.muscleape.excel.metadata;

import java.util.HashMap;
import java.util.Map;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Excel基础模型
 *
 * @author Muscleape
 */
@Data
public class BaseRowModel {

    /**
     * 每列样式
     */
    private Map<Integer, CellStyle> cellStyleMap = new HashMap<>();

    public void addStyle(Integer row, CellStyle cellStyle) {
        cellStyleMap.put(row, cellStyle);
    }

    public CellStyle getStyle(Integer row) {
        return cellStyleMap.get(row);
    }
}
