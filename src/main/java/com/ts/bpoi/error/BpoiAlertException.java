package com.ts.bpoi.error;

import com.ts.bpoi.base.BpoiConstants;

/**
 * 提示性的异常（运行时异常）
 * @author Bob
 */
public class BpoiAlertException extends BpoiException {

    private String code;

    public BpoiAlertException(String message) {
        super(message);
        this.code = BpoiConstants.commonReturnStatus.FAIL.getValue();
    }

    public BpoiAlertException(String code, String message) {
        super(message);
        this.code = code;
    }
    
    @Override
    public String getCode() {
        return code;
    }

}
