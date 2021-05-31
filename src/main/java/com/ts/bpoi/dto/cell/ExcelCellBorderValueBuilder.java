package com.ts.bpoi.dto.cell;

import com.ts.bpoi.dto.ExcelCellDTO;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel单元格Builder（带边框的数据）
 * @author Bob
 */
public class ExcelCellBorderValueBuilder implements ExcelCellBuilder {

    private ExcelCellDTO excelCellDTO;

    public ExcelCellBorderValueBuilder(int relativeRow, int column, String value) {
        ExcelCellDTO excelCellDTO = new ExcelCellDTO().setValue(value).setRelativeRow(relativeRow).setColumn(column);
        this.excelCellDTO = excelCellDTO;
    }

    /**
     * 构建单元格
     * @return
     */
    @Override
    public ExcelCellDTO buildCell() {
        return excelCellDTO.setHorizontal(HorizontalAlignment.CENTER).setAllBorder(BorderStyle.THIN);
    }

}
