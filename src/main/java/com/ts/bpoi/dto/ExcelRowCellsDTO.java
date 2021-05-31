package com.ts.bpoi.dto;

import java.util.List;

/**
 * Excel某行的单元格
 * @Author Bob
 */
public class ExcelRowCellsDTO {

    private int row;                            // 单元格所在行

    private int maxColumn;                      // 该行的最大列数

    private List<ExcelCellDTO> cellList;        // 该行的单元格列表

    public ExcelRowCellsDTO(int row, List<ExcelCellDTO> cellList) {
        this.row = row;
        this.cellList = cellList;
        maxColumn = 0;
        if (cellList != null && cellList.size() > 0) {
            for (ExcelCellDTO cell : cellList) {
                if (cell.getColumn() > maxColumn) {
                    maxColumn = cell.getColumn();
                }
            }
        }
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getMaxColumn() {
        return maxColumn;
    }

    public void setMaxColumn(int maxColumn) {
        this.maxColumn = maxColumn;
    }

    public List<ExcelCellDTO> getCellList() {
        return cellList;
    }

    public void setCellList(List<ExcelCellDTO> cellList) {
        this.cellList = cellList;
    }
}
