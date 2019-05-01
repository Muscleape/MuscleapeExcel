package com.muscleape.excel.analysis;

import com.muscleape.excel.metadata.MuscleapeSheet;

import java.util.List;

/**
 * Excel file analyser
 *
 * @author Muscleape
 */
public interface ExcelAnalyser {

    /**
     * parse one sheet
     *
     * @param sheetParam
     */
    void analysis(MuscleapeSheet sheetParam);

    /**
     * parse all sheets
     */
    void analysis();

    /**
     * get all sheet of workbook
     *
     * @return all sheets
     */
    List<MuscleapeSheet> getSheets();

}
