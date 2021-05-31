package com.ts.bpoi.dto.cell;

import com.ts.bpoi.base.BpoiConstants;
import com.ts.bpoi.dto.ExcelCellDTO;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * Excel单元格Builder（带边框的Key）
 * @author Bob
 */
public class ExcelCellBorderKeyBuilder implements ExcelCellBuilder {

    private ExcelCellDTO excelCellDTO;

    public ExcelCellBorderKeyBuilder(int relativeRow, int column, String value) {
        ExcelCellDTO excelCellDTO = new ExcelCellDTO().setRelativeRow(relativeRow).setColumn(column).setValue(value);
        this.excelCellDTO = excelCellDTO;
    }

    /**
     * 构建单元格
     * @return
     */
    @Override
    public ExcelCellDTO buildCell() {
        return excelCellDTO.setHorizontal(HorizontalAlignment.CENTER).setBackgroundColor(BpoiConstants.EXCEL_THEME_COLOR)
                .setAllBorder(BorderStyle.THIN);
    }

}
