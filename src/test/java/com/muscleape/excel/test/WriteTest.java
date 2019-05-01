package com.muscleape.excel.test;

import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.test.listen.AfterWriteHandlerImpl;
import com.muscleape.excel.test.model.WriteModel;
import com.muscleape.excel.test.util.FileUtil;
import com.muscleape.excel.MuscleapeExcelFactory;
import com.muscleape.excel.ExcelWriter;
import com.muscleape.excel.metadata.MuscleapeTable;
import com.muscleape.excel.support.ExcelTypeEnum;
import org.junit.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

import static com.muscleape.excel.test.util.DataUtil.*;

public class WriteTest {

    @Test
    public void writeV2007() throws IOException {
        OutputStream out = new FileOutputStream("D:\\Desktop\\临时文件\\2007.xlsx");
        ExcelWriter writer = MuscleapeExcelFactory.getWriter(out);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        MuscleapeSheet sheet1 = new MuscleapeSheet(1, 3);
        sheet1.setSheetName("第一个sheet");

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0, 10000);
        columnWidth.put(1, 40000);
        columnWidth.put(2, 10000);
        columnWidth.put(3, 10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        MuscleapeSheet sheet2 = new MuscleapeSheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        //writer.write1(null, sheet2);
        writer.write(createTestListJavaMode(), sheet2);
        //需要合并单元格
        writer.merge(5, 20, 1, 1);

        //写第三个sheet包含多个table情况
        MuscleapeSheet sheet3 = new MuscleapeSheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        MuscleapeTable table1 = new MuscleapeTable(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        MuscleapeTable table2 = new MuscleapeTable(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }


    @Test
    public void writeV2007WithTemplate() throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("temp.xlsx");
        OutputStream out = new FileOutputStream("D:\\Desktop\\临时文件\\2007.xlsx");
        ExcelWriter writer = MuscleapeExcelFactory.getWriterWithTemp(inputStream, out, ExcelTypeEnum.XLSX, true);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        MuscleapeSheet sheet1 = new MuscleapeSheet(1, 3);
        sheet1.setSheetName("第一个sheet");
        sheet1.setStartRow(20);

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0, 10000);
        columnWidth.put(1, 40000);
        columnWidth.put(2, 10000);
        columnWidth.put(3, 10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        MuscleapeSheet sheet2 = new MuscleapeSheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        sheet2.setStartRow(20);
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        MuscleapeSheet sheet3 = new MuscleapeSheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        sheet3.setStartRow(30);
        MuscleapeTable table1 = new MuscleapeTable(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        MuscleapeTable table2 = new MuscleapeTable(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }

    @Test
    public void writeV2007WithTemplateAndHandler() throws IOException {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("temp.xlsx");
        OutputStream out = new FileOutputStream("D:\\Desktop\\临时文件\\2007.xlsx");
        ExcelWriter writer = MuscleapeExcelFactory.getWriterWithTempAndHandler(inputStream, out, ExcelTypeEnum.XLSX, true,
                new AfterWriteHandlerImpl());
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        MuscleapeSheet sheet1 = new MuscleapeSheet(1, 3);
        sheet1.setSheetName("第一个sheet");
        sheet1.setStartRow(20);

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0, 10000);
        columnWidth.put(1, 40000);
        columnWidth.put(2, 10000);
        columnWidth.put(3, 10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        MuscleapeSheet sheet2 = new MuscleapeSheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        sheet2.setStartRow(20);
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        MuscleapeSheet sheet3 = new MuscleapeSheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        sheet3.setStartRow(30);
        MuscleapeTable table1 = new MuscleapeTable(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        MuscleapeTable table2 = new MuscleapeTable(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();

    }


    @Test
    public void writeV2003() throws IOException {
        OutputStream out = new FileOutputStream("D:\\Desktop\\临时文件\\2003.xls");
        ExcelWriter writer = MuscleapeExcelFactory.getWriter(out, ExcelTypeEnum.XLS, true);
        //写第一个sheet, sheet1  数据全是List<String> 无模型映射关系
        MuscleapeSheet sheet1 = new MuscleapeSheet(1, 3);
        sheet1.setSheetName("第一个sheet");

        //设置列宽 设置每列的宽度
        Map columnWidth = new HashMap();
        columnWidth.put(0, 10000);
        columnWidth.put(1, 40000);
        columnWidth.put(2, 10000);
        columnWidth.put(3, 10000);
        sheet1.setColumnWidthMap(columnWidth);
        sheet1.setHead(createTestListStringHead());
        //or 设置自适应宽度
        //sheet1.setAutoWidth(Boolean.TRUE);
        writer.write1(createTestListObject(), sheet1);

        //写第二个sheet sheet2  模型上打有表头的注解，合并单元格
        MuscleapeSheet sheet2 = new MuscleapeSheet(2, 3, WriteModel.class, "第二个sheet", null);
        sheet2.setTableStyle(createTableStyle());
        writer.write(createTestListJavaMode(), sheet2);

        //写第三个sheet包含多个table情况
        MuscleapeSheet sheet3 = new MuscleapeSheet(3, 0);
        sheet3.setSheetName("第三个sheet");
        MuscleapeTable table1 = new MuscleapeTable(1);
        table1.setHead(createTestListStringHead());
        writer.write1(createTestListObject(), sheet3, table1);

        //写sheet2  模型上打有表头的注解
        MuscleapeTable table2 = new MuscleapeTable(2);
        table2.setTableStyle(createTableStyle());
        table2.setClazz(WriteModel.class);
        writer.write(createTestListJavaMode(), sheet3, table2);

        writer.finish();
        out.close();
    }
}
