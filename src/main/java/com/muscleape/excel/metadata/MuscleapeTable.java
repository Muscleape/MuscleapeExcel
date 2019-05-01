package com.muscleape.excel.metadata;

import lombok.Data;

import java.util.List;

/**
 * @author Muscleape
 */
@Data
public class MuscleapeTable {
    /**
     *
     */
    private Class<? extends BaseRowModel> clazz;

    /**
     *
     */
    private List<List<String>> head;

    /**
     *
     */
    private int tableNo;

    /**
     *
     */
    private MuscleapeTableStyle tableStyle;

    public MuscleapeTable(Integer tableNo) {
        this.tableNo = tableNo;
    }
}
