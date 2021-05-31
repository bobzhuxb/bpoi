package com.ts.bpoi.dto;

import java.util.StringJoiner;

/**
 * Excel导入SAX解析的参数
 * @author Bob
 */
public class ExcelSaxImportParam {

    private StringJoiner errInfoSj;     // 错误信息

    private Integer maxErrHintCount;    // 最大错误提示数

    private Integer errHintCount;       // 已累计错误数

    public ExcelSaxImportParam() {}

    public ExcelSaxImportParam(StringJoiner errInfoSj, Integer maxErrHintCount, Integer errHintCount) {
        this.errInfoSj = errInfoSj;
        this.maxErrHintCount = maxErrHintCount;
        this.errHintCount = errHintCount;
    }

    public StringJoiner getErrInfoSj() {
        return errInfoSj;
    }

    public void setErrInfoSj(StringJoiner errInfoSj) {
        this.errInfoSj = errInfoSj;
    }

    public Integer getMaxErrHintCount() {
        return maxErrHintCount;
    }

    public void setMaxErrHintCount(Integer maxErrHintCount) {
        this.maxErrHintCount = maxErrHintCount;
    }

    public Integer getErrHintCount() {
        return errHintCount;
    }

    public void setErrHintCount(Integer errHintCount) {
        this.errHintCount = errHintCount;
    }
}
