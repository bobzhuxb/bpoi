package com.ts.bpoi.dto;

/**
 * Excel标题单元格配置DTO
 * @author Bob
 */
public class ExcelTitleDTO implements Comparable<ExcelTitleDTO> {

    private String titleName;       // 标题名（字段或Key）
    private String titleContent;    // 标题内容
    private Integer colIndex;       // 列序号（从1开始）
    
    public ExcelTitleDTO() {}
    
    public ExcelTitleDTO(String titleName, String titleContent) {
        this.titleName = titleName;
        this.titleContent = titleContent;
    }

    public ExcelTitleDTO(Integer colIndex, String titleName, String titleContent) {
        this.colIndex = colIndex;
        this.titleName = titleName;
        this.titleContent = titleContent;
    }

    public String getTitleName() {
        return titleName;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public String getTitleContent() {
        return titleContent;
    }

    public void setTitleContent(String titleContent) {
        this.titleContent = titleContent;
    }

    public Integer getColIndex() {
        return colIndex;
    }

    public void setColIndex(Integer colIndex) {
        this.colIndex = colIndex;
    }

    @Override
    public int compareTo(ExcelTitleDTO o) {
        if (this.colIndex == null || o.colIndex == null) {
            return 0;
        }
        return this.colIndex - o.colIndex;
    }
}
