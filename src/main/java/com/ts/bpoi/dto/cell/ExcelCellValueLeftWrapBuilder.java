package com.ts.bpoi.dto.cell;

import com.ts.bpoi.dto.ExcelCellDTO;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Excel单元格Builder（带边框的数据，居左，可换行）
 * @author Bob
 */
public class ExcelCellValueLeftWrapBuilder implements ExcelCellBuilder {

    private ExcelCellDTO excelCellDTO;

    public ExcelCellValueLeftWrapBuilder() {
        ExcelCellDTO excelCellDTO = new ExcelCellDTO();
        this.excelCellDTO = excelCellDTO;
    }

    /**
     * 构建单元格
     * @return
     */
    @Override
    public ExcelCellDTO buildCell() {
        return excelCellDTO.setHorizontal(HorizontalAlignment.LEFT)
                .setVertical(VerticalAlignment.CENTER).setAllBorder(BorderStyle.THIN).setWrapText(true);
    }

}
