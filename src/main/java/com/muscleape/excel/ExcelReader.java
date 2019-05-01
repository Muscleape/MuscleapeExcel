package com.muscleape.excel;

import com.muscleape.excel.analysis.ExcelAnalyser;
import com.muscleape.excel.analysis.ExcelAnalyserImpl;
import com.muscleape.excel.context.AnalysisContext;
import com.muscleape.excel.event.AnalysisEventListener;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.support.ExcelTypeEnum;

import java.io.InputStream;
import java.util.List;

/**
 * Excel readers are all read in event mode.
 *
 * @author Muscleape
 */
public class ExcelReader {

    /**
     * Analyser
     */
    private ExcelAnalyser analyser;

    /**
     * Create new reader
     *
     * @param in            the POI filesystem that contains the Workbook stream
     * @param customContent {@link AnalysisEventListener#invoke(Object, AnalysisContext) }AnalysisContext
     * @param eventListener Callback method after each row is parsed
     */
    public ExcelReader(InputStream in, Object customContent,
                       AnalysisEventListener eventListener) {
        this(in, customContent, eventListener, true);
    }

    /**
     * Create new reader
     *
     * @param in
     * @param customContent {@link AnalysisEventListener#invoke(Object, AnalysisContext) }AnalysisContext
     * @param eventListener
     * @param trim          The content of the form is empty and needs to be empty. The purpose is to be fault-tolerant,
     *                      because there are often table contents with spaces that can not be converted into custom
     *                      types. For example: '1234 ' contain a space cannot be converted to int.
     */
    public ExcelReader(InputStream in, Object customContent,
                       AnalysisEventListener eventListener, boolean trim) {
        ExcelTypeEnum excelTypeEnum = ExcelTypeEnum.valueOf(in);
        validateParam(in, eventListener);
        analyser = new ExcelAnalyserImpl(in, excelTypeEnum, customContent, eventListener, trim);
    }

    /**
     * Parse all sheet content by default
     */
    public void read() {
        analyser.analysis();
    }

    /**
     * Parse the specified sheet，SheetNo start from 1
     *
     * @param sheet Read sheet
     */
    public void read(MuscleapeSheet sheet) {
        analyser.analysis(sheet);
    }


    /**
     * Parse the workBook get all sheets
     *
     * @return workBook all sheets
     */
    public List<MuscleapeSheet> getSheets() {
        return analyser.getSheets();
    }

    /**
     * validate param
     *
     * @param in
     * @param eventListener
     */
    private void validateParam(InputStream in, AnalysisEventListener eventListener) {
        if (eventListener == null) {
            throw new IllegalArgumentException("AnalysisEventListener can not null");
        } else if (in == null) {
            throw new IllegalArgumentException("InputStream can not null");
        }
    }
}
