package com.ts.bpoi.dto.cell;

import com.ts.bpoi.dto.ExcelCellDTO;

/**
 * Excel单元格Builder
 * @author Bob
 */
public interface ExcelCellBuilder {

    /**
     * 构建单元格
     * @return
     */
    ExcelCellDTO buildCell();

}
