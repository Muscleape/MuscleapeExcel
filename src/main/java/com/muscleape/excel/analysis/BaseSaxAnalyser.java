package com.muscleape.excel.analysis;

import com.muscleape.excel.context.AnalysisContext;
import com.muscleape.excel.event.AnalysisEventListener;
import com.muscleape.excel.event.AnalysisEventRegisterCenter;
import com.muscleape.excel.event.OneRowAnalysisFinishEvent;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.util.TypeUtil;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Muscleape
 */
public abstract class BaseSaxAnalyser implements AnalysisEventRegisterCenter, ExcelAnalyser {

    protected AnalysisContext analysisContext;

    private LinkedHashMap<String, AnalysisEventListener> listeners = new LinkedHashMap<String, AnalysisEventListener>();

    /**
     * execute method
     */
    protected abstract void execute();


    @Override
    public void appendLister(String name, AnalysisEventListener listener) {
        if (!listeners.containsKey(name)) {
            listeners.put(name, listener);
        }
    }

    @Override
    public void analysis(MuscleapeSheet sheetParam) {
        execute();
    }

    @Override
    public void analysis() {
        execute();
    }

    /**
     *
     */
    @Override
    public void cleanAllListeners() {
        listeners = new LinkedHashMap<>();
    }

    @Override
    public void notifyListeners(OneRowAnalysisFinishEvent event) {
        analysisContext.setCurrentRowAnalysisResult(event.getData());
        /** Parsing header content **/
        if (analysisContext.getCurrentRowNum() < analysisContext.getCurrentSheet().getHeadLineMun()) {
            if (analysisContext.getCurrentRowNum() <= analysisContext.getCurrentSheet().getHeadLineMun() - 1) {
                analysisContext.buildExcelHeadProperty(null, (List<String>) analysisContext.getCurrentRowAnalysisResult());
            }
        } else {
            List<String> content = converter((List<String>) event.getData());
            /** Parsing Analyze the body content **/
            analysisContext.setCurrentRowAnalysisResult(content);
            if (listeners.size() == 1) {
                analysisContext.setCurrentRowAnalysisResult(content);
            }
            /**  notify all event listeners **/
            for (Map.Entry<String, AnalysisEventListener> entry : listeners.entrySet()) {
                entry.getValue().invoke(analysisContext.getCurrentRowAnalysisResult(), analysisContext);
            }
        }
    }

    private List<String> converter(List<String> data) {
        List<String> list = new ArrayList<>();
        if (data != null) {
            for (String str : data) {
                list.add(TypeUtil.formatFloat(str));
            }
        }
        return list;
    }

}
