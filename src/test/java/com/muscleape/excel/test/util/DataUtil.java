package com.muscleape.excel.test.util;

import com.muscleape.excel.metadata.MuscleapeFont;
import com.muscleape.excel.metadata.MuscleapeTableStyle;
import com.muscleape.excel.test.model.MuscleapeWriteModel;
import com.muscleape.excel.test.model.MuscleapeWriteModelOldVersion;
import com.muscleape.excel.test.model.WriteModel;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DataUtil {


    public static List<List<Object>> createTestListObject() {
        List<List<Object>> object = new ArrayList<List<Object>>();
        for (int i = 0; i < 1000; i++) {
            List<Object> da = new ArrayList<Object>();
            da.add("字符串" + i);
            da.add(Long.valueOf(187837834l + i));
            da.add(Integer.valueOf(2233 + i));
            da.add(Double.valueOf(2233.00 + i));
            da.add(Float.valueOf(2233.0f + i));
            da.add(new Date());
            da.add(new BigDecimal("3434343433554545" + i));
            da.add(Short.valueOf((short) i));
            object.add(da);
        }
        return object;
    }

    /**
     * 按照模型组装数据，测试字段类型处理
     *
     * @return java.util.List<com.muscleape.excel.test.model.MuscleapeWriteModel>
     * @author Muscleape
     * @date 2019/8/23 16:47
     */
    public static List<MuscleapeWriteModel> createMuscleapeTestListObject() {
        List<MuscleapeWriteModel> modelList = new ArrayList<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateStr = "2019-08-23 17:59:00";
        // 2019-08-23 17:59:00 对应的时间戳
        Long dateLong = 1566554340000L;
        try {
            MuscleapeWriteModel modelOne = MuscleapeWriteModel.builder()
                    .id(1)
                    .str("我的名字")
                    .date(dateFormat.parse(dateStr))
                    .dateStr(dateStr)
                    .timestampLong(dateLong)
                    .timestampString("1566554340000")
                    .intKeyValue(1)
                    .intStrKeyValue(1)
                    .strKeyValue("1")
                    .emptyKeyValue("1")
                    .otherKeyValueNoDefault("6")
                    .otherKeyValueWithDefault("6")
                    .nullKeyValueNoDefault(null)
                    .nullKeyValueWithDefault(null)
                    .build();
            modelList.add(modelOne);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return modelList;
    }

    public static List<MuscleapeWriteModelOldVersion> createMuscleapeTestListObjectOldVersion() {
        List<MuscleapeWriteModelOldVersion> modelList = new ArrayList<>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateStr = "2019-08-23 17:59:00";
        // 2019-08-23 17:59:00 对应的时间戳
        Long dateLong = 1566554340000L;
        try {
            MuscleapeWriteModelOldVersion modelOne = MuscleapeWriteModelOldVersion.builder()
                    .id(1)
                    .str("我的名字")
                    .date(dateFormat.parse(dateStr))
                    .dateStr(dateStr)
                    .timestampLong(dateLong)
                    .intKeyValue(1)
                    .intStrKeyValue(1)
                    .strKeyValue("1")
                    .emptyKeyValue("1")
                    .build();
            modelList.add(modelOne);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return modelList;
    }

    public static List<List<String>> createTestListStringHead() {
        //写sheet3  模型上没有注解，表头数据动态传入
        List<List<String>> head = new ArrayList<List<String>>();
        List<String> headCoulumn1 = new ArrayList<String>();
        List<String> headCoulumn2 = new ArrayList<String>();
        List<String> headCoulumn3 = new ArrayList<String>();
        List<String> headCoulumn4 = new ArrayList<String>();
        List<String> headCoulumn5 = new ArrayList<String>();

        headCoulumn1.add("第一列");
        headCoulumn1.add("第一列");
        headCoulumn1.add("第一列");
        headCoulumn2.add("第一列");
        headCoulumn2.add("第一列");
        headCoulumn2.add("第一列");

        headCoulumn3.add("第二列");
        headCoulumn3.add("第二列");
        headCoulumn3.add("第二列");
        headCoulumn4.add("第三列");
        headCoulumn4.add("第三列2");
        headCoulumn4.add("第三列2");
        headCoulumn5.add("第一列");
        headCoulumn5.add("第3列");
        headCoulumn5.add("第4列");

        head.add(headCoulumn1);
        head.add(headCoulumn2);
        head.add(headCoulumn3);
        head.add(headCoulumn4);
        head.add(headCoulumn5);
        return head;
    }

    public static List<WriteModel> createTestListJavaMode() {
        List<WriteModel> model1s = new ArrayList<WriteModel>();
        for (int i = 0; i < 1000; i++) {
            WriteModel model1 = new WriteModel();
            model1.setP1("第一列，第行");
            model1.setP2("121212jjj");
            model1.setP3(33 + i);
            model1.setP4(44);
            model1.setP5("555");
            model1.setP6(666.2f);
            model1.setP7(new BigDecimal("454545656343434" + i));
            model1.setP8(new Date());
            model1.setP9("llll9999>&&&&&6666^^^^");
            model1.setP10(1111.77 + i);
            model1.setP11(i % 2);
            model1s.add(model1);
        }
        return model1s;
    }

    public static MuscleapeTableStyle createTableStyle() {
        MuscleapeTableStyle tableStyle = new MuscleapeTableStyle();
        MuscleapeFont headFont = new MuscleapeFont();
        headFont.setBold(true);
        headFont.setFontHeightInPoints((short) 22);
        headFont.setFontName("楷体");
        tableStyle.setTableHeadFont(headFont);
        tableStyle.setTableHeadBackGroundColor(IndexedColors.BLUE);

        MuscleapeFont contentFont = new MuscleapeFont();
        contentFont.setBold(true);
        contentFont.setFontHeightInPoints((short) 22);
        contentFont.setFontName("黑体");
        tableStyle.setTableContentFont(contentFont);
        tableStyle.setTableContentBackGroundColor(IndexedColors.GREEN);
        return tableStyle;
    }
}
