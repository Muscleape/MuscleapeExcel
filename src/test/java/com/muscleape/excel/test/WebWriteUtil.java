package com.muscleape.excel.test;

import com.muscleape.excel.ExcelWriter;
import com.muscleape.excel.MuscleapeExcelFactory;
import com.muscleape.excel.exception.ExcelGenerateException;
import com.muscleape.excel.metadata.BaseRowModel;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.support.ExcelTypeEnum;
import org.springframework.util.ObjectUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Author: Muscleape
 * @Date: 2019-05-01
 * @Description: WEB中生成Excel并下载的使用方式(可以直接放到WEB项目中 ， 作为Excel工具类使用)
 */

public class WebWriteUtil {
    /**
     * 导出 Excel ：一个 sheet，带表头
     *
     * @param response  HttpServletResponse
     * @param list      数据 list，每个元素为一个 BaseRowModel
     * @param fileName  导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param object    映射实体类，Excel 模型
     */
    public static void writeExcel(HttpServletResponse response, List<? extends BaseRowModel> list,
                                  String fileName, String sheetName, BaseRowModel object) {
        ExcelWriter excelWriter = MuscleapeExcelFactory.getWriter(getOutputStream(fileName + ExcelTypeEnum.XLSX.getValue(), response), ExcelTypeEnum.XLSX, true);
        MuscleapeSheet sheet = new MuscleapeSheet(1, 0, object.getClass());
        sheet.setSheetName(sheetName);
        excelWriter.write(list, sheet);
        excelWriter.finish();
    }

    /**
     * 导出 Excel：入参List<List<Model>>，每一个内层list导出到一个sheet页
     *
     * @param response  HttpServletResponse
     * @param fileName  导出的文件名(不带文件类型后缀)
     * @param sheetName 导出文件中的 sheet 名
     * @return void
     * @author Muscleape
     * @date 2019/4/27 15:28
     */
    public static void writeExcelInSheetsWithList(HttpServletResponse response, List<List<? extends BaseRowModel>> baseRowModelsList,
                                                  String fileName, String sheetName) {
        ExcelWriter excelWriter = MuscleapeExcelFactory.getWriter(getOutputStream(fileName + ExcelTypeEnum.XLSX.getValue(), response), ExcelTypeEnum.XLSX, true);
        int pageNum = 1;
        for (List<? extends BaseRowModel> baseRowModels : baseRowModelsList) {
            if (!ObjectUtils.isEmpty(baseRowModels)) {
                MuscleapeSheet sheet = new MuscleapeSheet(pageNum, 0, baseRowModels.get(0).getClass());
                sheet.setSheetName(sheetName + "(" + pageNum + ")");
                excelWriter.write(baseRowModels, sheet);
                pageNum++;
            }
        }
        excelWriter.finish();
    }

    /**
     * 导出 Excel：入参List<List<Model>>，每一个内层list导出到一个sheet页
     *
     * @param response  HttpServletResponse
     * @param fileName  导出的文件名(不带文件类型后缀)
     * @param sheetName 导出文件中的 sheet 名
     * @param object    映射实体类，Excel 模型，继承BaseRowModel
     * @return void
     * @author Muscleape
     * @date 2019/4/27 15:28
     */
    public static void writeExcelInSheetsWithList(HttpServletResponse response, List<List<? extends BaseRowModel>> baseRowModelsList,
                                                  String fileName, String sheetName,
                                                  BaseRowModel object) {
        ExcelWriter excelWriter = MuscleapeExcelFactory.getWriter(getOutputStream(fileName + ExcelTypeEnum.XLSX.getValue(), response), ExcelTypeEnum.XLSX, true);
        int pageNum = 1;
        for (List<? extends BaseRowModel> baseRowModels : baseRowModelsList) {
            MuscleapeSheet sheet = new MuscleapeSheet(pageNum, 0, object.getClass());
            sheet.setSheetName(sheetName + "(" + pageNum + ")");
            excelWriter.write(baseRowModels, sheet);
            pageNum++;
        }
        excelWriter.finish();
    }

    /**
     * 导出 Excel ：多个 sheet，带表头，默认每个sheet页1000条数据
     *
     * @param response  HttpServletResponse
     * @param list      数据 list，每个元素为一个 BaseRowModel
     * @param fileName  导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param object    映射实体类，Excel 模型
     */
    public static void writeExcelWithSheets(HttpServletResponse response, List<? extends BaseRowModel> list,
                                            String fileName, String sheetName, BaseRowModel object) {
        writeExcelWithSheets(response, list, 1000, fileName, sheetName, object);
    }

    /**
     * 导出 Excel ：多个 sheet，带表头
     *
     * @param response  HttpServletResponse
     * @param list      数据 list，每个元素为一个 BaseRowModel
     * @param size      每个sheet页中数据条数
     * @param fileName  导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param object    映射实体类，Excel 模型
     */
    public static void writeExcelWithSheets(HttpServletResponse response, List<? extends BaseRowModel> list,
                                            int size, String fileName, String sheetName, BaseRowModel object) {
        ExcelWriter excelWriter = MuscleapeExcelFactory.getWriter(getOutputStream(fileName + ExcelTypeEnum.XLSX.getValue(), response), ExcelTypeEnum.XLSX, true);
        size = (0 >= size) ? 1000 : size;
        int sheetCount = (list.size() % size == 0) ? (list.size() / size) : (list.size() / size + 1);
        for (int i = 1; i <= sheetCount; i++) {
            MuscleapeSheet sheet = new MuscleapeSheet(i, 0, object.getClass());
            sheet.setSheetName(sheetName + "(" + i + ")");
            List tempList = list.stream().skip((i - 1) * size).limit(size).collect(Collectors.toList());
            excelWriter.write(tempList, sheet);
        }
        excelWriter.finish();
    }

    /**
     * 导出文件时为Writer生成OutputStream
     */
    private static OutputStream getOutputStream(String fileName, HttpServletResponse response) {
        try {
            fileName = new String(fileName.getBytes(), "ISO-8859-1");
            response.addHeader("Content-Disposition", "filename=" + fileName);
            return response.getOutputStream();
        } catch (IOException e) {
            throw new ExcelGenerateException("创建文件失败！");
        }
    }

}
