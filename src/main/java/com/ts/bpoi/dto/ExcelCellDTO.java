package com.ts.bpoi.dto;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excel单元格配置DTO
 * @author Bob
 */
public class ExcelCellDTO {

    /**
     * 位置
     */
    private Integer relativeRow;            // 单元格所在相对行（实际数据行前后分别从0开始计）
    private Integer column;                 // 单元格所在列（从0开始计）
    /**
     * 样式
     */
    private HorizontalAlignment horizontal; // 水平对齐方式
    private VerticalAlignment vertical;     // 垂直对齐方式
    private Short backgroundColor;          // 背景色
    private BorderStyle borderTop;          // 上边框
    private BorderStyle borderBottom;       // 下边框
    private BorderStyle borderLeft;         // 左边框
    private BorderStyle borderRight;        // 右边框
    private String fontName;                // 字体名称
    private Short fontSize;                 // 字体大小
    private Boolean fontItalic;             // 是否斜体
    private Boolean fontBold;               // 是否粗体
    private Short fontColor;                // 字体颜色
    private Boolean wrapText;               // 是否换行
    /**
     * 数据
     */
    private String value;        // 单元格数据

    public Integer getRelativeRow() {
        return relativeRow;
    }

    public ExcelCellDTO setRelativeRow(Integer relativeRow) {
        this.relativeRow = relativeRow;
        return this;
    }

    public Integer getColumn() {
        return column;
    }

    public ExcelCellDTO setColumn(Integer column) {
        this.column = column;
        return this;
    }

    public HorizontalAlignment getHorizontal() {
        return horizontal;
    }

    public ExcelCellDTO setHorizontal(HorizontalAlignment horizontal) {
        this.horizontal = horizontal;
        return this;
    }

    public VerticalAlignment getVertical() {
        return vertical;
    }

    public ExcelCellDTO setVertical(VerticalAlignment vertical) {
        this.vertical = vertical;
        return this;
    }

    public Short getBackgroundColor() {
        return backgroundColor;
    }

    public ExcelCellDTO setBackgroundColor(Short backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public ExcelCellDTO setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
        return this;
    }

    public BorderStyle getBorderBottom() {
        return borderBottom;
    }

    public ExcelCellDTO setBorderBottom(BorderStyle borderBottom) {
        this.borderBottom = borderBottom;
        return this;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public ExcelCellDTO setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
        return this;
    }

    public BorderStyle getBorderRight() {
        return borderRight;
    }

    public ExcelCellDTO setBorderRight(BorderStyle borderRight) {
        this.borderRight = borderRight;
        return this;
    }

    public String getFontName() {
        return fontName;
    }

    public ExcelCellDTO setFontName(String fontName) {
        this.fontName = fontName;
        return this;
    }

    public Short getFontSize() {
        return fontSize;
    }

    public ExcelCellDTO setFontSize(Short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public Boolean getFontItalic() {
        return fontItalic;
    }

    public ExcelCellDTO setFontItalic(Boolean fontItalic) {
        this.fontItalic = fontItalic;
        return this;
    }

    public Boolean getFontBold() {
        return fontBold;
    }

    public ExcelCellDTO setFontBold(Boolean fontBold) {
        this.fontBold = fontBold;
        return this;
    }

    public Short getFontColor() {
        return fontColor;
    }

    public ExcelCellDTO setFontColor(Short fontColor) {
        this.fontColor = fontColor;
        return this;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public ExcelCellDTO setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
        return this;
    }

    public String getValue() {
        return value;
    }

    public ExcelCellDTO setValue(String value) {
        this.value = value;
        return this;
    }

    /**
     * 设置所有框线
     * @param borderTop
     * @param borderBottom
     * @param borderLeft
     * @param borderRight
     * @return
     */
    public ExcelCellDTO setAllBorder(BorderStyle borderTop, BorderStyle borderBottom,
                                     BorderStyle borderLeft, BorderStyle borderRight) {
        this.borderTop = borderTop;
        this.borderBottom = borderBottom;
        this.borderLeft = borderLeft;
        this.borderRight = borderRight;
        return this;
    }

    /**
     * 设置所有框线
     * @param border
     * @return
     */
    public ExcelCellDTO setAllBorder(BorderStyle border) {
        setAllBorder(border, border, border, border);
        return this;
    }
}
