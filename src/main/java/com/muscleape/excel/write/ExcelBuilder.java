package com.muscleape.excel.write;

import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.metadata.MuscleapeTable;

import java.util.List;

/**
 * @author Muscleape
 */
public interface ExcelBuilder {

    /**
     * workBook increase data
     *
     * @param data     java basic type or java model extend BaseModel
     * @param startRow Start row number
     */
    void addContent(List data, int startRow);

    /**
     * WorkBook increase data
     *
     * @param data       java basic type or java model extend BaseModel
     * @param sheetParam Write the sheet
     */
    void addContent(List data, MuscleapeSheet sheetParam);

    /**
     * WorkBook increase data
     *
     * @param data       java basic type or java model extend BaseModel
     * @param sheetParam Write the sheet
     * @param table      Write the table
     */
    void addContent(List data, MuscleapeSheet sheetParam, MuscleapeTable table);

    /**
     * Creates new cell range. Indexes are zero-based.
     *
     * @param firstRow Index of first row
     * @param lastRow  Index of last row (inclusive), must be equal to or larger than {@code firstRow}
     * @param firstCol Index of first column
     * @param lastCol  Index of last column (inclusive), must be equal to or larger than {@code firstCol}
     */
    void merge(int firstRow, int lastRow, int firstCol, int lastCol);

    /**
     * Close io
     */
    void finish();
}
