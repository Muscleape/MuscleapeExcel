package com.muscleape.excel.test.model;

import com.muscleape.excel.annotation.ExcelProperty;
import com.muscleape.excel.metadata.BaseRowModel;
import lombok.Data;

@Data
public class BaseReadModel extends BaseRowModel {
    @ExcelProperty(index = 0)
    protected String str;

    @ExcelProperty(index = 1)
    protected Float ff;
}
