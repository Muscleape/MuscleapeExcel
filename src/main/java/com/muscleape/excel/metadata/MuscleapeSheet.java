package com.muscleape.excel.metadata;

import lombok.Data;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Muscleape
 */
@Data
public class MuscleapeSheet {

    /**
     *
     */
    private int headLineMun;

    /**
     * Starting from 1
     */
    private int sheetNo;

    /**
     *
     */
    private String sheetName;

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
    private MuscleapeTableStyle tableStyle;

    /**
     * column with
     */
    private Map<Integer, Integer> columnWidthMap = new HashMap<>();

    /**
     *
     */
    private Boolean autoWidth = Boolean.FALSE;

    /**
     *
     */
    private int startRow = 0;


    public MuscleapeSheet(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public MuscleapeSheet(int sheetNo, int headLineMun) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
    }

    public MuscleapeSheet(int sheetNo, int headLineMun, Class<? extends BaseRowModel> clazz) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
        this.clazz = clazz;
    }

    public MuscleapeSheet(int sheetNo, int headLineMun, Class<? extends BaseRowModel> clazz, String sheetName,
                          List<List<String>> head) {
        this.sheetNo = sheetNo;
        this.clazz = clazz;
        this.headLineMun = headLineMun;
        this.sheetName = sheetName;
        this.head = head;
    }

    public void setClazz(Class<? extends BaseRowModel> clazz) {
        this.clazz = clazz;
        if (headLineMun == 0) {
            this.headLineMun = 1;
        }
    }

    @Override
    public String toString() {
        return "MuscleapeSheet{" +
                "headLineMun=" + headLineMun +
                ", sheetNo=" + sheetNo +
                ", sheetName='" + sheetName + '\'' +
                ", clazz=" + clazz +
                ", head=" + head +
                ", tableStyle=" + tableStyle +
                ", columnWidthMap=" + columnWidthMap +
                '}';
    }
}
