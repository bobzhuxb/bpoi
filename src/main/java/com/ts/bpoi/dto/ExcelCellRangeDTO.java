package com.ts.bpoi.dto;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Excel单元格范围配置DTO
 * @author Bob
 */
public class ExcelCellRangeDTO {

    private int fromRow;                    // 起始行（从0开始）
    private int toRow;                      // 终止行
    private int fromColumn;                 // 起始列（从0开始）
    private int toColumn;                   // 终止列

    private BorderStyle borderTop;          // 上边框
    private BorderStyle borderBottom;       // 下边框
    private BorderStyle borderLeft;         // 左边框
    private BorderStyle borderRight;        // 右边框

    public ExcelCellRangeDTO() {

    }

    public ExcelCellRangeDTO(int fromRow, int toRow, int fromColumn, int toColumn) {
        this.fromRow = fromRow;
        this.toRow = toRow;
        this.fromColumn = fromColumn;
        this.toColumn = toColumn;
    }

    public int getFromRow() {
        return fromRow;
    }

    public ExcelCellRangeDTO setFromRow(int fromRow) {
        this.fromRow = fromRow;
        return this;
    }

    public int getToRow() {
        return toRow;
    }

    public ExcelCellRangeDTO setToRow(int toRow) {
        this.toRow = toRow;
        return this;
    }

    public int getFromColumn() {
        return fromColumn;
    }

    public ExcelCellRangeDTO setFromColumn(int fromColumn) {
        this.fromColumn = fromColumn;
        return this;
    }

    public int getToColumn() {
        return toColumn;
    }

    public ExcelCellRangeDTO setToColumn(int toColumn) {
        this.toColumn = toColumn;
        return this;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public ExcelCellRangeDTO setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
        return this;
    }

    public BorderStyle getBorderBottom() {
        return borderBottom;
    }

    public ExcelCellRangeDTO setBorderBottom(BorderStyle borderBottom) {
        this.borderBottom = borderBottom;
        return this;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public ExcelCellRangeDTO setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
        return this;
    }

    public BorderStyle getBorderRight() {
        return borderRight;
    }

    public ExcelCellRangeDTO setBorderRight(BorderStyle borderRight) {
        this.borderRight = borderRight;
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
    public ExcelCellRangeDTO setAllBorder(BorderStyle borderTop, BorderStyle borderBottom,
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
    public ExcelCellRangeDTO setAllBorder(BorderStyle border) {
        setAllBorder(border, border, border, border);
        return this;
    }
}
