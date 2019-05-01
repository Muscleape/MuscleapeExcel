package com.muscleape.excel.event;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Muscleape
 */
@Data
public class OneRowAnalysisFinishEvent {

    public OneRowAnalysisFinishEvent(Object content) {
        this.data = content;
    }

    public OneRowAnalysisFinishEvent(String[] content, int length) {
        if (content != null) {
            List<String> ls = new ArrayList<>(length);
            for (int i = 0; i <= length; i++) {
                ls.add(content[i]);
            }
            data = ls;
        }
    }

    private Object data;

}
