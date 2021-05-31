package com.ts.bpoi.util;

import com.ts.bpoi.dto.ExcelSaxReadCellDTO;

import java.util.List;

/**
 * SAX方式解析Excel时读取到的行
 * @author Bob
 */
public interface IExcelSaxRowRead {

	/**
	 * 获取返回数据
	 * @return
	 */
	<T> List<T> getList();

	/**
	 * 解析数据
	 * @param index 行号
	 * @param dataList 该行的数据
	 */
	void parse(int index, List<ExcelSaxReadCellDTO> dataList);

}
