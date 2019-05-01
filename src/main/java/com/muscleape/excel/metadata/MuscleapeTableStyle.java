package com.muscleape.excel.metadata;

import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author Muscleape
 */
@Data
public class MuscleapeTableStyle {

    /**
     * 表头背景颜色
     */
    private IndexedColors tableHeadBackGroundColor;

    /**
     * 表头字体样式
     */
    private MuscleapeFont tableHeadFont;

    /**
     * 表格内容字体样式
     */
    private MuscleapeFont tableContentFont;

    /**
     * 表格内容背景颜色
     */
    private IndexedColors tableContentBackGroundColor;
}
