package com.muscleape.excel.metadata;

import lombok.Data;

/**
 * @author Muscleape
 */
@Data
public class MuscleapeIndexValue {

    private String v_index;
    private String v_value;

    public MuscleapeIndexValue(String v_index, String v_value) {
        super();
        this.v_index = v_index;
        this.v_value = v_value;
    }

    @Override
    public String toString() {
        return "MuscleapeIndexValue [v_index=" + v_index + ", v_value=" + v_value + "]";
    }
}
