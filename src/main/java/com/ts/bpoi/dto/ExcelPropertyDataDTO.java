package com.ts.bpoi.dto;

import java.util.regex.Pattern;

/**
 * Excel解析后的相关属性
 * @author Bob
 */
public class ExcelPropertyDataDTO {

    private String propertyName;    // 属性名

    private String value;           // 表头名称

    private Integer index;          // 排序

    private String regex;           // 验证正则

    private String emptyToDefault;  // （该值为非空字符串时有效，且忽略nullable的影响）

    private Boolean nullable;       // 是否允许空值

    private String dateFormat;      // 日期格式

    private String[] optionList;    // 取值范围

    private Pattern pattern;        // 验证正则的Pattern

    public ExcelPropertyDataDTO() {
        this.index = 0;
        this.nullable = true;
    }

    public String getPropertyName() {
        return propertyName;
    }

    public void setPropertyName(String propertyName) {
        this.propertyName = propertyName;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getRegex() {
        return regex;
    }

    public void setRegex(String regex) {
        this.regex = regex;
    }

    public String getEmptyToDefault() {
        return emptyToDefault;
    }

    public void setEmptyToDefault(String emptyToDefault) {
        this.emptyToDefault = emptyToDefault;
    }

    public Boolean getNullable() {
        return nullable;
    }

    public void setNullable(Boolean nullable) {
        this.nullable = nullable;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public String[] getOptionList() {
        return optionList;
    }

    public void setOptionList(String[] optionList) {
        this.optionList = optionList;
    }

    public Pattern getPattern() {
        return pattern;
    }

    public void setPattern(Pattern pattern) {
        this.pattern = pattern;
    }
}
