package com.muscleape.excel.analysis;

import com.muscleape.excel.analysis.v03.XlsSaxAnalyser;
import com.muscleape.excel.analysis.v07.XlsxSaxAnalyser;
import com.muscleape.excel.context.AnalysisContext;
import com.muscleape.excel.context.AnalysisContextImpl;
import com.muscleape.excel.event.AnalysisEventListener;
import com.muscleape.excel.exception.ExcelAnalysisException;
import com.muscleape.excel.metadata.MuscleapeSheet;
import com.muscleape.excel.modelbuild.ModelBuildEventListener;
import com.muscleape.excel.support.ExcelTypeEnum;

import java.io.InputStream;
import java.util.List;

/**
 * @author Muscleape
 */
public class ExcelAnalyserImpl implements ExcelAnalyser {

    private AnalysisContext analysisContext;

    private BaseSaxAnalyser saxAnalyser;

    public ExcelAnalyserImpl(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
                             AnalysisEventListener eventListener, boolean trim) {
        analysisContext = new AnalysisContextImpl(inputStream, excelTypeEnum, custom, eventListener, trim);
    }

    private BaseSaxAnalyser getSaxAnalyser() {
        if (saxAnalyser != null) {
            return this.saxAnalyser;
        }
        try {
            if (analysisContext.getExcelType() != null) {
                switch (analysisContext.getExcelType()) {
                    case XLS:
                        this.saxAnalyser = new XlsSaxAnalyser(analysisContext);
                        break;
                    case XLSX:
                        this.saxAnalyser = new XlsxSaxAnalyser(analysisContext);
                        break;
                }
            } else {
                try {
                    this.saxAnalyser = new XlsxSaxAnalyser(analysisContext);
                } catch (Exception e) {
                    if (!analysisContext.getInputStream().markSupported()) {
                        throw new ExcelAnalysisException(
                                "Xls must be available markSupported,you can do like this <code> new "
                                        + "BufferedInputStream(new FileInputStream(\"/xxxx\"))</code> ");
                    }
                    this.saxAnalyser = new XlsSaxAnalyser(analysisContext);
                }
            }
        } catch (Exception e) {
            throw new ExcelAnalysisException("File type error，io must be available markSupported,you can do like "
                    + "this <code> new BufferedInputStream(new FileInputStream(\\\"/xxxx\\\"))</code> \"", e);
        }
        return this.saxAnalyser;
    }

    @Override
    public void analysis(MuscleapeSheet sheetParam) {
        analysisContext.setCurrentSheet(sheetParam);
        analysis();
    }

    @Override
    public void analysis() {
        BaseSaxAnalyser saxAnalyser = getSaxAnalyser();
        appendListeners(saxAnalyser);
        saxAnalyser.execute();
        analysisContext.getEventListener().doAfterAllAnalysed(analysisContext);
    }

    @Override
    public List<MuscleapeSheet> getSheets() {
        BaseSaxAnalyser saxAnalyser = getSaxAnalyser();
        saxAnalyser.cleanAllListeners();
        return saxAnalyser.getSheets();
    }

    private void appendListeners(BaseSaxAnalyser saxAnalyser) {
        saxAnalyser.cleanAllListeners();
        if (analysisContext.getCurrentSheet() != null && analysisContext.getCurrentSheet().getClazz() != null) {
            saxAnalyser.appendLister("model_build_listener", new ModelBuildEventListener());
        }
        if (analysisContext.getEventListener() != null) {
            saxAnalyser.appendLister("user_define_listener", analysisContext.getEventListener());
        }
    }

}
