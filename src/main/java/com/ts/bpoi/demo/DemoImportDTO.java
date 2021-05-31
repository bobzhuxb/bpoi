package com.ts.bpoi.demo;

import com.ts.bpoi.annotation.ExcelProperty;

/**
 * 测试导入DTO类
 * @author Bob
 */
public class DemoImportDTO {

    @ExcelProperty(value = "序号", index = 1, nullable = false)
    private String index;

    @ExcelProperty(value = "名称", index = 2, nullable = false)
    private String name;

    @ExcelProperty(value = "内容", index = 3, nullable = false, regex = "^.{1,30}$")
    private String content;

    public String getIndex() {
        return index;
    }

    public void setIndex(String index) {
        this.index = index;
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

    @Override
    public String toString() {
        return "DemoImportDTO{" +
                "index='" + index + '\'' +
                ", name='" + name + '\'' +
                ", content='" + content + '\'' +
                '}';
    }
}
