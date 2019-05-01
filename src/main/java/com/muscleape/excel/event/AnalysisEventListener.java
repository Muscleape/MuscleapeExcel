package com.muscleape.excel.event;

import com.muscleape.excel.context.AnalysisContext;

/**
 *
 *
 * @author Muscleape
 */
public abstract class AnalysisEventListener<T> {

    /**
     * when analysis one row trigger invoke function
     *
     * @param object  one row data
     * @param context analysis context
     */
    public abstract void invoke(T object, AnalysisContext context);

    /**
     * if have something to do after all  analysis
     *
     * @param context
     */
    public abstract void doAfterAllAnalysed(AnalysisContext context);
}
