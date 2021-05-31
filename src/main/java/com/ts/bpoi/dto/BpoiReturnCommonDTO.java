package com.ts.bpoi.dto;

import com.ts.bpoi.base.BpoiConstants;

/**
 * 数据返回的共通类
 * @author Bob
 */
public class BpoiReturnCommonDTO<T> {

    // 返回状态
    private String resultCode;

    // 返回消息
    private String errMsg;

    // 返回数据
    private T data;

    public BpoiReturnCommonDTO() {
        this(BpoiConstants.commonReturnStatus.SUCCESS.getValue());
    }

    public BpoiReturnCommonDTO(String resultCode) {
        this.resultCode = resultCode;
    }

    public BpoiReturnCommonDTO(String resultCode, String errMsg) {
        this(resultCode);
        this.errMsg = errMsg;
    }

    public BpoiReturnCommonDTO(String resultCode, String errMsg, T data) {
        this(resultCode, errMsg);
        this.data = data;
    }

    public BpoiReturnCommonDTO(T data) {
        this(BpoiConstants.commonReturnStatus.SUCCESS.getValue(), null, data);
    }

    public static BpoiReturnCommonDTO commonErrorReturn(String errMsg) {
        return new BpoiReturnCommonDTO<>(BpoiConstants.commonReturnStatus.FAIL.getValue(), errMsg);
    }

    public String getResultCode() {
        return resultCode;
    }

    public void setResultCode(String resultCode) {
        this.resultCode = resultCode;
    }

    public String getErrMsg() {
        return errMsg;
    }

    public void setErrMsg(String errMsg) {
        this.errMsg = errMsg;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }
}
