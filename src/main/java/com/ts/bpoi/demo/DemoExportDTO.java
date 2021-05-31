package com.ts.bpoi.demo;

import com.ts.bpoi.annotation.ExcelProperty;

/**
 * 测试导出DTO类
 * @author Bob
 */
public class DemoExportDTO {

    @ExcelProperty(value = "名称", index = 5)
    private String name;

    @ExcelProperty(value = "内容", index = 2)
    private String content;

    @ExcelProperty(value = "描述", index = 2)
    private String descr;

    @ExcelProperty(value = "备注", index = 3)
    private String memo;

    public DemoExportDTO(String name, String content, String descr, String memo) {
        this.name = name;
        this.content = content;
        this.descr = descr;
        this.memo = memo;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getDescr() {
        return descr;
    }

    public void setDescr(String descr) {
        this.descr = descr;
    }

    public String getMemo() {
        return memo;
    }

    public void setMemo(String memo) {
        this.memo = memo;
    }
}
