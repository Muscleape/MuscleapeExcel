package com.muscleape.excel.metadata;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * @author Muscleape
 */
@Data
@AllArgsConstructor
public class MuscleapeCellRange {
    private int firstRow;
    private int lastRow;
    private int firstCol;
    private int lastCol;
}
