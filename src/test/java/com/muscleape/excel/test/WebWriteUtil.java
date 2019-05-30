package com.muscleape.excel.test;

import com.muscleape.excel.ExcelWriter;
import com.muscleape.excel.MuscleapeExcelFactory;
import com.muscleape.excel.exception.ExcelGenerateException;
import com.muscleape.excel.metadata.BaseRowModel;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.support.ExcelTypeEnum;
import org.springframework.util.ObjectUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

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
     * 多次查询拼装成一个List后压缩导出
     *
     * @param response
     * @param lists
     * @param fileName
     * @param object
     * @return void
     * @author Muscleape
     * @date 2019/5/29 10:57
     */
    public static void writeLists2ExcelFileInZip(HttpServletResponse response, List<List<? extends BaseRowModel>> lists,
                                                 String fileName, BaseRowModel object) throws IOException {
        String pathName = mkdirPath(fileName);
        fileName = pathName + "/" + pathName;

        int count = 1;
        for (List<? extends BaseRowModel> list : lists) {
            // 创建导出文件
            OutputStream outputStream = new FileOutputStream(fileName + "(" + count + ").xlsx");
            ExcelWriter excelWriter = MuscleapeExcelFactory.getWriter(outputStream);
            // 创建sheet
            MuscleapeSheet sheet = new MuscleapeSheet(1, 0, object.getClass());
            sheet.setSheetName(String.valueOf(count));
            excelWriter.write(list, sheet);
            excelWriter.finish();
            count += 1;
        }
        lists.clear();

        writeZipOutputStream(response, lists.size(), pathName, fileName);
    }

    /**
     * 创建临时excel文件路径
     *
     * @param fileName
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/5/29 10:25
     */
    private static String mkdirPath(String fileName) {
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS");
        String localDateTime = LocalDateTime.now().format(dateTimeFormatter);
        String pathName = fileName + "_" + localDateTime + "_" + Thread.currentThread().getId();
        File path = new File(pathName);
        if (!path.exists()) {
            path.mkdirs();
        }
        return pathName;
    }

    /**
     * 写压缩包文件
     *
     * @param response
     * @param fileCount
     * @param pathName
     * @param fileName
     * @return void
     * @author Muscleape
     * @date 2019/5/29 10:38
     */
    private static void writeZipOutputStream(HttpServletResponse response, int fileCount, String pathName, String fileName) {
        //设置压缩流：直接写入response，实现边压缩边下载
        ZipOutputStream zipOutputStream = createZipOutputStream(response, pathName);
        //循环将文件写入压缩流
        DataOutputStream dataOutputStream = null;
        for (int i = 1; i <= fileCount; i++) {
            String tempFileName = fileName + "(" + i + ").xlsx";
            File file = new File(tempFileName);
            try {
                //添加ZipEntry，并ZipEntry中写入文件流
                zipOutputStream.putNextEntry(new ZipEntry(tempFileName));
                dataOutputStream = new DataOutputStream(zipOutputStream);
                InputStream inputStream = new FileInputStream(file);
                byte[] b = new byte[100];
                int length;
                while ((length = inputStream.read(b)) != -1) {
                    dataOutputStream.write(b, 0, length);
                }
                inputStream.close();
                zipOutputStream.closeEntry();
            } catch (IOException e) {
                e.printStackTrace();
                return;
            } finally {
                // 删除产生的Excel文件
                // file.delete();
            }
        }
        try {
            // 关闭流
            dataOutputStream.flush();
            dataOutputStream.close();
            zipOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }
    }

    private static ZipOutputStream createZipOutputStream(HttpServletResponse response, String zipFileName) {
        try {
            //设置压缩包的名字
            String downloadName = new String((zipFileName + ".zip").getBytes(), "ISO-8859-1");
            response.setHeader("Content-Disposition", "attachment;fileName=\"" + downloadName + "\"");

            //设置压缩流：直接写入response，实现边压缩边下载
            ZipOutputStream zipOutputStream = new ZipOutputStream(new BufferedOutputStream(response.getOutputStream()));
            //设置压缩方法
            zipOutputStream.setMethod(ZipOutputStream.DEFLATED);
            return zipOutputStream;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
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
