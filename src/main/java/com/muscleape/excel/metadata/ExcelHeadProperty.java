package com.muscleape.excel.metadata;

import com.muscleape.excel.annotation.ExcelColumnNum;
import com.muscleape.excel.annotation.ExcelProperty;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Define the header attribute of excel
 *
 * @author Muscleape
 */
@Data
public class ExcelHeadProperty {

    /**
     * Custom class
     */
    private Class<? extends BaseRowModel> headClazz;

    /**
     * A two-dimensional array describing the header
     */
    private List<List<String>> head;

    /**
     * Attributes described by the header
     */
    private List<ExcelColumnProperty> columnPropertyList = new ArrayList<>();

    /**
     * Attributes described by the header
     */
    private Map<Integer, ExcelColumnProperty> excelColumnPropertyMap1 = new HashMap<>();

    public ExcelHeadProperty(Class<? extends BaseRowModel> headClazz, List<List<String>> head) {
        this.headClazz = headClazz;
        this.head = head;
        initColumnProperties();
    }

    /**
     *
     */
    private void initColumnProperties() {
        if (this.headClazz != null) {
            List<Field> fieldList = new ArrayList<>();
            Class tempClass = this.headClazz;
            //When the parent class is null, it indicates that the parent class (Object class) has reached the top
            // level.
            while (tempClass != null) {
                fieldList.addAll(Arrays.asList(tempClass.getDeclaredFields()));
                //Get the parent class and give it to yourself
                tempClass = tempClass.getSuperclass();
            }
            List<List<String>> headList = new ArrayList<>();
            for (Field f : fieldList) {
                initOneColumnProperty(f);
            }
            //对列排序
            Collections.sort(columnPropertyList);
            if (head == null || head.size() == 0) {
                for (ExcelColumnProperty excelColumnProperty : columnPropertyList) {
                    headList.add(excelColumnProperty.getHead());
                }
                this.head = headList;
            }
        }
    }

    /**
     * @param f
     */
    private void initOneColumnProperty(Field f) {
        ExcelProperty p = f.getAnnotation(ExcelProperty.class);
        ExcelColumnProperty excelHeadProperty = null;
        if (p != null) {
            excelHeadProperty = new ExcelColumnProperty();
            excelHeadProperty.setField(f);
            excelHeadProperty.setHead(Arrays.asList(p.value()));
            excelHeadProperty.setIndex(p.index());
            excelHeadProperty.setFormat(p.format());
            excelHeadProperty.setKeyValue(p.keyValue());
            excelColumnPropertyMap1.put(p.index(), excelHeadProperty);
        } else {
            ExcelColumnNum columnNum = f.getAnnotation(ExcelColumnNum.class);
            if (columnNum != null) {
                excelHeadProperty = new ExcelColumnProperty();
                excelHeadProperty.setField(f);
                excelHeadProperty.setIndex(columnNum.value());
                excelHeadProperty.setFormat(columnNum.format());
                excelHeadProperty.setKeyValue(columnNum.keyValue());
                excelColumnPropertyMap1.put(columnNum.value(), excelHeadProperty);
            }
        }
        if (excelHeadProperty != null) {
            this.columnPropertyList.add(excelHeadProperty);
        }
    }

    /**
     *
     */
    public void appendOneRow(List<String> row) {

        for (int i = 0; i < row.size(); i++) {
            List<String> list;
            if (head.size() <= i) {
                list = new ArrayList<>();
                head.add(list);
            } else {
                list = head.get(0);
            }
            list.add(row.get(i));
        }
    }

    /**
     * @param columnNum
     * @return
     */
    public ExcelColumnProperty getExcelColumnProperty(int columnNum) {
        return excelColumnPropertyMap1.get(columnNum);
    }

    /**
     * Calculate all cells that need to be merged
     *
     * @return cells that need to be merged
     */
    public List<MuscleapeCellRange> getCellRangeModels() {
        List<MuscleapeCellRange> cellRanges = new ArrayList<>();
        for (int i = 0; i < head.size(); i++) {
            List<String> columnValues = head.get(i);
            for (int j = 0; j < columnValues.size(); j++) {
                int lastRow = getLastRangNum(j, columnValues.get(j), columnValues);
                int lastColumn = getLastRangNum(i, columnValues.get(j), getHeadByRowNum(j));
                boolean flag = (lastRow > j || lastColumn > i) && lastRow >= 0 && lastColumn >= 0;
                if (flag) {
                    cellRanges.add(new MuscleapeCellRange(j, lastRow, i, lastColumn));
                }
            }
        }
        return cellRanges;
    }

    public List<String> getHeadByRowNum(int rowNum) {
        List<String> l = new ArrayList<>(head.size());
        for (List<String> list : head) {
            if (list.size() > rowNum) {
                l.add(list.get(rowNum));
            } else {
                l.add(list.get(list.size() - 1));
            }
        }
        return l;
    }

    /**
     * Get the last consecutive string position
     *
     * @param j      current value position
     * @param value  value content
     * @param values values
     * @return the last consecutive string position
     */
    private int getLastRangNum(int j, String value, List<String> values) {
        if (value == null) {
            return -1;
        }
        if (j > 0) {
            String preValue = values.get(j - 1);
            if (value.equals(preValue)) {
                return -1;
            }
        }
        int last = j;
        for (int i = last + 1; i < values.size(); i++) {
            String current = values.get(i);
            if (value.equals(current)) {
                last = i;
            } else {
                // if i>j && !value.equals(current) Indicates that the continuous range is exceeded
                if (i > j) {
                    break;
                }
            }
        }
        return last;
    }

    public int getRowNum() {
        int headRowNum = 0;
        for (List<String> list : head) {
            if (list != null && list.size() > 0) {
                if (list.size() > headRowNum) {
                    headRowNum = list.size();
                }
            }
        }
        return headRowNum;
    }
}
