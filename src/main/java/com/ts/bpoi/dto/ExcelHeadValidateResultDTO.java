package com.ts.bpoi.dto;

import java.util.Map;

/**
 * Excel表头解析结果
 * @author Bob
 */
public class ExcelHeadValidateResultDTO {

    private Map<Integer, ExcelPropertyDataDTO> headMap;     // 解析出的结果（Key：从0开始计的列号 Value：解析出的列信息）

    private Integer maxColumn;              // 表头最大列号（从1开始计）

    public Map<Integer, ExcelPropertyDataDTO> getHeadMap() {
        return headMap;
    }

    public void setHeadMap(Map<Integer, ExcelPropertyDataDTO> headMap) {
        this.headMap = headMap;
    }

    public Integer getMaxColumn() {
        return maxColumn;
    }

    public void setMaxColumn(Integer maxColumn) {
        this.maxColumn = maxColumn;
    }
}
