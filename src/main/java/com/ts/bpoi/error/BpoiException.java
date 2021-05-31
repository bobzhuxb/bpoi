package com.ts.bpoi.error;

import com.ts.bpoi.base.BpoiConstants;

/**
 * 共通异常类（运行时异常）
 * @author Bob
 */
public class BpoiException extends RuntimeException {

    /**
     * 错误代码
     */
    private String code;

    /**
     * 错误明细，用于记录日志
     */
    private String errDetail;

    // 另：继承的message字段用于返回给前端提示给用户错误信息

    public BpoiException(String message) {
        super(message);
        this.code = BpoiConstants.commonReturnStatus.FAIL.getValue();
    }

    public BpoiException(String code, String message) {
        super(message);
        this.code = code;
    }

    public BpoiException(String message, Throwable e) {
        super(message, e);
        this.code = BpoiConstants.commonReturnStatus.FAIL.getValue();
    }

    public static BpoiException errWithDetail(String message, String errDetail, Throwable e) {
        BpoiException commonException = new BpoiException(message, e);
        commonException.setErrDetail(errDetail);
        return commonException;
    }
    
    public String getCode() {
        return code;
    }

    public String getErrDetail() {
        return errDetail;
    }

    public void setErrDetail(String errDetail) {
        this.errDetail = errDetail;
    }
}
