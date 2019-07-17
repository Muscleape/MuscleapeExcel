package com.muscleape.excel.test;

import com.muscleape.excel.ExcelWriter;
import com.muscleape.excel.MuscleapeExcelFactory;
import com.muscleape.excel.exception.ExcelGenerateException;
import com.muscleape.excel.metadata.BaseRowModel;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.support.ExcelTypeEnum;
import lombok.extern.slf4j.Slf4j;
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
 * @Date: 2019-04-16
 * @Description:
 */
// @Slf4j
public class MuscleapeExcelUtil {

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
     * @author wangzhongqi
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
     * @author wangzhongqi
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
     * 导出 Excel ：多个 sheet，带表头，默认每个sheet页10000条数据
     *
     * @param response  HttpServletResponse
     * @param list      数据 list，每个元素为一个 BaseRowModel
     * @param fileName  导出的文件名
     * @param sheetName 导入文件的 sheet 名
     * @param object    映射实体类，Excel 模型
     */
    public static void writeExcelWithSheets(HttpServletResponse response, List<? extends BaseRowModel> list,
                                            String fileName, String sheetName, BaseRowModel object) {
        writeExcelWithSheets(response, list, 10000, fileName, sheetName, object);
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
        size = (0 >= size) ? 10000 : size;
        int sheetCount = (list.size() % size == 0) ? (list.size() / size) : (list.size() / size + 1);
        for (int i = 1; i <= sheetCount; i++) {
            MuscleapeSheet sheet = new MuscleapeSheet(i, 0, object.getClass());
            sheet.setSheetName(sheetName + "(" + i + ")");
            List tempList = list.stream().skip((i - 1) * size).limit(size).collect(Collectors.toList());
            excelWriter.write(tempList, sheet);
            tempList.clear();
        }
        excelWriter.finish();
        list.clear();
    }


    /**
     * 导出多个Excel文件，压缩ZIP包后下载
     *
     * @param response
     * @param list     导出数据列表
     * @param size     每个Excel文件中数据条数(最小10000条数据)
     * @param fileName 文件名
     * @return void
     * @author Muscleape
     * @date 2019/5/10 17:46
     */
    public static void writeOneList2ExcelFileInZip(HttpServletResponse response, List<? extends BaseRowModel> list,
                                                   int size, String fileName, BaseRowModel object) throws IOException {
        writeListExcelFileInZip(response, list, size, fileName, object, false);
    }

    /**
     * 导出本地ZIP文件，并返回ZIP文件路径
     *
     * @param response
     * @param list
     * @param size
     * @param fileName
     * @param object
     * @return java.lang.String
     * @author Muscleape
     * @date 2019/7/8 10:27
     */
    public static String writeOneList2ExcelFileInZipReturnPath(HttpServletResponse response, List<? extends BaseRowModel> list,
                                                               int size, String fileName, BaseRowModel object) {
        String pathName = mkdirPath(fileName);
        // fileName = pathName + File.separator + fileName;

        int fileCount = createExcelFile(size, list, pathName, fileName, object);

        // 导出处理
        String zipFilePath = writeZipOutputStream(response, fileCount, pathName, fileName, true);

        return zipFilePath;
    }

    /**
     * 生成Excel文件，并返回文件个数
     *
     * @param size
     * @param list
     * @param pathName
     * @param fileName
     * @param object
     * @return void
     * @author Muscleape
     * @date 2019/7/8 10:33
     */
    private static int createExcelFile(int size, List<? extends BaseRowModel> list, String pathName, String fileName, BaseRowModel object) {
        size = (0 >= size) ? 10000 : size;
        int fileCount = (list.size() % size == 0) ? (list.size() / size) : (list.size() / size + 1);

        try {
            for (int i = 1; i <= fileCount; i++) {
                // 创建导出文件
                OutputStream outputStream = null;
                outputStream = new FileOutputStream(pathName + File.separator + fileName + "(" + i + ").xlsx");
                ExcelWriter excelWriter = MuscleapeExcelFactory.getWriter(outputStream);
                List<? extends BaseRowModel> tempList = list.stream().skip((i - 1) * size).limit(size).collect(Collectors.toList());
                // 创建sheet
                MuscleapeSheet sheet = new MuscleapeSheet(1, 0, object.getClass());
                sheet.setSheetName(String.valueOf(i));
                excelWriter.write(tempList, sheet);
                excelWriter.finish();
                tempList.clear();
            }
            list.clear();
            return fileCount;
        } catch (FileNotFoundException e) {
            // log.error("[MuscleapeError] create excel file error!");
            e.printStackTrace();
            return 0;
        }
    }

    /**
     * ZIP方式导出文件
     *
     * @param response
     * @param list
     * @param size
     * @param fileName
     * @param object
     * @param isLocal
     * @return void
     * @author Muscleape
     * @date 2019/7/8 10:25
     */
    public static void writeListExcelFileInZip(HttpServletResponse response, List<? extends BaseRowModel> list,
                                               int size, String fileName, BaseRowModel object, boolean isLocal) throws IOException {
        String pathName = mkdirPath(fileName);
        // fileName = pathName + File.separator + fileName;

        int fileCount = createExcelFile(size, list, pathName, fileName, object);

        // 导出处理
        writeZipOutputStream(response, fileCount, pathName, fileName, isLocal);
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
    public static void writeLists2ExcelFileInZip(HttpServletResponse response, List<List<? extends
            BaseRowModel>> lists,
                                                 String fileName, BaseRowModel object) throws IOException {
        String pathName = mkdirPath(fileName);
        fileName = pathName + File.separator + pathName;

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

        writeZipOutputStream(response, lists.size(), pathName, fileName, false);
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
        String pathName = "exportExcelFile" + File.separator + fileName + "_" + localDateTime + "_" + Thread.currentThread().getId();
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
     * @param isLocal   压缩文件流类型（true:本地生成ZIP文件;false:生成zip文件的web下载流）
     */
    private static String writeZipOutputStream(HttpServletResponse response, int fileCount, String pathName, String fileName, boolean isLocal) {

        ZipOutputStream zipOutputStream;
        if (isLocal) {
            // 生成zip文件，将Excel文件写入到zip中
            zipOutputStream = createZipOutputStream4Local(pathName, fileName);
        } else {
            //设置压缩流：直接写入response，实现边压缩边下载
            zipOutputStream = createZipOutputStream4WebDownload(response, fileName);
        }
        if (ObjectUtils.isEmpty(zipOutputStream)) {
            return "";
        }

        //循环将文件写入压缩流
        DataOutputStream dataOutputStream = null;
        for (int i = 1; i <= fileCount; i++) {
            String tempFileName = pathName + File.separator + fileName + "(" + i + ").xlsx";
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
                // log.info("[MuscleapeInfo] ZipEntry add excel file success!");
            } catch (IOException e) {
                // log.error("[MuscleapeError] ZipEntry add excel file error!");
                e.printStackTrace();
                return "";
            }
        }
        // try {
        //     // 删除产生的Excel文件及生成的目录
        //     Files.walk(Paths.get(pathName)).sorted(Comparator.reverseOrder()).map(Path::toFile).peek(System.out::println).forEach(File::delete);
        // } catch (IOException e) {
        //     log.error("[MuscleapeError] delete the excel file and the path error!");
        //     e.printStackTrace();
        //     return;
        // }
        try {
            // 关闭流
            dataOutputStream.flush();
            dataOutputStream.close();
            zipOutputStream.close();
            // log.info("[MuscleapeInfo] close the file stream success!");
        } catch (IOException e) {
            // log.error("[MuscleapeError] close the file stream error!");
            e.printStackTrace();
            return "";
        }
        // 返回生成的zip文件路径
        return pathName + File.separator + fileName + ".zip";
    }

    /**
     * 生成web下载使用的压缩文件流
     *
     * @param response
     * @param zipFileName
     * @return java.util.zip.ZipOutputStream
     * @author Muscleape
     * @date 2019/7/7 15:02
     */
    private static ZipOutputStream createZipOutputStream4WebDownload(HttpServletResponse response, String zipFileName) {
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
            // log.error("[MuscleapeError] create ZipOutputStream for web download file error!");
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 生成本地本地zip文件的压缩文件流
     *
     * @param zipFilePath
     * @param zipFileName
     * @return java.util.zip.ZipOutputStream
     * @author Muscleape
     * @date 2019/7/7 15:03
     */
    private static ZipOutputStream createZipOutputStream4Local(String zipFilePath, String zipFileName) {
        try {
            if (!new File(zipFilePath).exists()) {
                new File(zipFilePath).mkdirs();
            }
            File zipFile = new File(zipFilePath + File.separator + zipFileName + ".zip");
            FileOutputStream fileOutputStream = new FileOutputStream(zipFile);
            ZipOutputStream zipOutputStream = new ZipOutputStream(new BufferedOutputStream(fileOutputStream));
            return zipOutputStream;
        } catch (Exception e) {
            // log.error("[MuscleapeError] create ZipOutputStream for local file error!");
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
