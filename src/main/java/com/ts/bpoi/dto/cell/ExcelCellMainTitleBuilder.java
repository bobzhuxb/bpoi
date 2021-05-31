package com.ts.bpoi.dto.cell;

import com.ts.bpoi.dto.ExcelCellDTO;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel单元格Builder（主标题）
 * @author Bob
 */
public class ExcelCellMainTitleBuilder implements ExcelCellBuilder {

    private ExcelCellDTO excelCellDTO;

    public ExcelCellMainTitleBuilder(int relativeRow, int column, String value) {
        ExcelCellDTO excelCellDTO = new ExcelCellDTO().setRelativeRow(relativeRow).setColumn(column).setValue(value);
        this.excelCellDTO = excelCellDTO;
    }

    /**
     * 构建单元格
     * @return
     */
    @Override
    public ExcelCellDTO buildCell() {
        return excelCellDTO.setHorizontal(HorizontalAlignment.CENTER);
    }

}
