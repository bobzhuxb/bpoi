package com.ts.bpoi.dto;

import com.ts.bpoi.base.BpoiConstants;

/**
 * SAX解析的Excel单元格值
 * @author Bob
 */
public class ExcelSaxReadCellDTO {

	private BpoiConstants.excelCellValueType cellType;		// 值类型

	private Object value;		// 值

	public ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType cellType, Object value) {
		this.cellType = cellType;
		this.value = value;
	}

	public BpoiConstants.excelCellValueType getCellType() {
		return cellType;
	}

	public void setCellType(BpoiConstants.excelCellValueType cellType) {
		this.cellType = cellType;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}

}
