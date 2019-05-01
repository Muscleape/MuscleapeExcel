package com.muscleape.excel.modelbuild;

import com.muscleape.excel.context.AnalysisContext;
import com.muscleape.excel.event.AnalysisEventListener;
import com.muscleape.excel.exception.ExcelGenerateException;
import com.muscleape.excel.metadata.ExcelHeadProperty;
import com.muscleape.excel.util.TypeUtil;
import net.sf.cglib.beans.BeanMap;
import org.jetbrains.annotations.NotNull;

import java.util.List;

/**
 * @author Muscleape
 */
public class ModelBuildEventListener extends AnalysisEventListener {

    @Override
    public void invoke(Object object, @NotNull AnalysisContext context) {
        if (context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz() != null) {
            try {
                Object resultModel = buildUserModel(context, (List<String>) object);
                context.setCurrentRowAnalysisResult(resultModel);
            } catch (Exception e) {
                throw new ExcelGenerateException(e);
            }
        }
    }

    private Object buildUserModel(@NotNull AnalysisContext context, List<String> stringList) throws Exception {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();
        Object resultModel = excelHeadProperty.getHeadClazz().newInstance();
        if (excelHeadProperty == null) {
            return resultModel;
        }
        BeanMap.create(resultModel).putAll(
                TypeUtil.getFieldValues(stringList, excelHeadProperty, context.use1904WindowDate()));
        return resultModel;
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
