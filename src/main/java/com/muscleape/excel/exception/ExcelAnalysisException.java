package com.muscleape.excel.exception;

import lombok.NoArgsConstructor;

/**
 * @author Muscleape
 */
@NoArgsConstructor
public class ExcelAnalysisException extends RuntimeException {

    public ExcelAnalysisException(String message) {
        super(message);
    }

    public ExcelAnalysisException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelAnalysisException(Throwable cause) {
        super(cause);
    }
}
