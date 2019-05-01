package com.muscleape.excel.test.model;

import com.muscleape.excel.annotation.ExcelProperty;
import lombok.Data;
import org.junit.validator.ValidateWith;

import java.math.BigDecimal;
import java.util.Date;

@Data
public class WriteModel extends BaseWriteModel {

    @ExcelProperty(value = {"表头3", "表头3", "表头3"}, index = 2)
    private int p3;

    @ExcelProperty(value = {"表头1", "表头4", "表头4"}, index = 3)
    private long p4;

    @ExcelProperty(value = {"表头5", "表头51", "表头52"}, index = 4)
    private String p5;

    @ExcelProperty(value = {"表头6", "表头61", "表头611"}, index = 5)
    private float p6;

    @ExcelProperty(value = {"表头6", "表头61", "表头612"}, index = 6)
    private BigDecimal p7;

    @ExcelProperty(value = {"表头6", "表头62", "表头621"}, index = 7)
    private Date p8;

    @ExcelProperty(value = {"表头6", "表头62", "表头622"}, index = 8)
    private String p9;

    @ExcelProperty(value = {"表头6", "表头62", "表头622"}, index = 9)
    private double p10;

    @ExcelProperty(value = {"表头7", "表头71", "性别"}, index = 10, keyValue = "{'0':'女','1':'男'}")
    private int p11;

    @Override
    public String toString() {
        return "JavaModel1{" +
                "p1='" + p1 + '\'' +
                ", p2='" + p2 + '\'' +
                ", p3=" + p3 +
                ", p4=" + p4 +
                ", p5='" + p5 + '\'' +
                ", p6=" + p6 +
                ", p7=" + p7 +
                ", p8=" + p8 +
                ", p9='" + p9 + '\'' +
                ", p10=" + p10 +
                ", p11=" + p11 +
                '}';
    }
}
