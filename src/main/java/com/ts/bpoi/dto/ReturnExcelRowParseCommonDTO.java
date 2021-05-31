package com.ts.bpoi.dto;

/**
 * Excel行解析返回的共通类
 * @author Bob
 */
public class ReturnExcelRowParseCommonDTO<T> extends BpoiReturnCommonDTO<T> {

    private Integer errHintCount;       // 解析出的错误计数

    public ReturnExcelRowParseCommonDTO(String code, String errMsg) {
        super(code, errMsg);
    }

    public ReturnExcelRowParseCommonDTO(T data, Integer errHintCount) {
        super(data);
        this.errHintCount = errHintCount;
    }

    public Integer getErrHintCount() {
        return errHintCount;
    }

    public void setErrHintCount(Integer errHintCount) {
        this.errHintCount = errHintCount;
    }
}
