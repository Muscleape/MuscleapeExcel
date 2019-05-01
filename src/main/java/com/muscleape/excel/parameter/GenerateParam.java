package com.muscleape.excel.parameter;

import com.muscleape.excel.support.ExcelTypeEnum;
import lombok.Data;

import java.io.OutputStream;

/**
 * @author Muscleape
 * @date 17/2/19
 * @return
 */
@Data
public class GenerateParam {

    private OutputStream outputStream;

    private String sheetName;

    private Class clazz;

    private ExcelTypeEnum type;

    public GenerateParam(String sheetName, Class clazz, OutputStream outputStream) {
        this.outputStream = outputStream;
        this.sheetName = sheetName;
        this.clazz = clazz;
    }
}
