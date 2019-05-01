package com.muscleape.excel.context;

import com.muscleape.excel.event.AnalysisEventListener;
import com.muscleape.excel.metadata.BaseRowModel;
import com.muscleape.excel.metadata.ExcelHeadProperty;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.support.ExcelTypeEnum;

import java.io.InputStream;
import java.util.List;

/**
 * A context is the main anchorage point of a excel reader.
 *
 * @author Muscleape
 */
public interface AnalysisContext {

    /**
     * Custom attribute
     */
    Object getCustom();

    /**
     * get current sheet
     *
     * @return current analysis sheet
     */
    MuscleapeSheet getCurrentSheet();

    /**
     * set current sheet
     *
     * @param sheet
     */
    void setCurrentSheet(MuscleapeSheet sheet);

    /**
     * get excel type
     *
     * @return excel type
     */
    ExcelTypeEnum getExcelType();

    /**
     * get in io
     *
     * @return file io
     */
    InputStream getInputStream();

    /**
     * custom listener
     *
     * @return listener
     */
    AnalysisEventListener getEventListener();

    /**
     * get current row
     *
     * @return
     */
    Integer getCurrentRowNum();

    /**
     * set  current row num
     *
     * @param row
     */
    void setCurrentRowNum(Integer row);

    /**
     * get total row ,Data may be inaccurate
     *
     * @param totalCount
     */
    void setTotalCount(Integer totalCount);

    /**
     * get excel head
     *
     * @return
     */
    ExcelHeadProperty getExcelHeadProperty();

    /**
     * @param clazz
     * @param headOneRow
     */
    void buildExcelHeadProperty(Class<? extends BaseRowModel> clazz, List<String> headOneRow);

    /**
     * if need to short match the content
     *
     * @return
     */
    boolean trim();

    /**
     * set current result
     *
     * @param result
     */
    void setCurrentRowAnalysisResult(Object result);


    /**
     * get current result
     *
     * @return get current result
     */
    Object getCurrentRowAnalysisResult();

    /**
     * Interrupt execution
     */
    void interrupt();

    /**
     * date use1904WindowDate
     *
     * @return
     */
    boolean use1904WindowDate();

    /**
     * date use1904WindowDate
     *
     * @param use1904WindowDate
     */
    void setUse1904WindowDate(boolean use1904WindowDate);
}
