package com.ts.bpoi.util;

import com.ts.bpoi.base.BpoiConstants;
import com.ts.bpoi.dto.*;
import com.ts.bpoi.error.BpoiAlertException;
import com.ts.bpoi.error.BpoiException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 当前行读取数据
 * @author Bob
 */
public class ExcelSaxRowRead implements IExcelSaxRowRead {

	private ExcelSaxImportParam param;		// 导入参数

	private Class<?> excelParseClass;		// 解析用Class

	private Map<Integer, ExcelPropertyDataDTO> headMap;     // 解析出的结果（Key：从0开始计的列号 Value：解析出的列信息）

	private Integer maxColumn;              // 表头最大列号（从1开始计）

	private IExcelReadRowHandler handler;	// 当前行的处理

	private List objList;					// 需要返回的数据

	public ExcelSaxRowRead(Class<?> excelParseClass, ExcelSaxImportParam param, IExcelReadRowHandler handler) {
		objList = new ArrayList();
		this.excelParseClass = excelParseClass;
		this.param = param;
		this.handler = handler;
	}

	@Override
	public <T> List<T> getList() {
		return objList;
	}

	/**
	 * 解析
	 * @param rowIndex 当前行号
	 * @param dataList 该行的数据
	 */
	@Override
	public void parse(int rowIndex, List<ExcelSaxReadCellDTO> dataList) {
		try {
			if (dataList == null || dataList.size() == 0) {
				return;
			}
			if (rowIndex < 1) {
				// 表头行的处理
				initHeadData(dataList);
			} else {
				// 数据行的处理
				addListData(rowIndex, dataList);
			}
		} catch (BpoiAlertException e) {
			throw e;
		} catch (Exception e) {
			throw new BpoiException(e.getMessage(), e);
		}
	}

	/**
	 * 集合元素处理（如果每一行单独处理，则不会返回最终解析出的列表）
	 * @param rowIndex 当前行号
	 * @param dataList 当前行的数据
	 */
	private void addListData(int rowIndex, List<ExcelSaxReadCellDTO> dataList) {
		int dataMaxCell = dataList.size();
		if (dataMaxCell > maxColumn) {
			param.getErrInfoSj().add("第" + (rowIndex + 1) + "行列数超过" + maxColumn);
			param.setErrHintCount(param.getErrHintCount() + 1);
			if (param.getMaxErrHintCount() != null && param.getErrHintCount() >= param.getMaxErrHintCount()) {
				throw new BpoiAlertException(param.getErrInfoSj().toString());
			}
		} else if (dataMaxCell < maxColumn) {
			// 数据行列数小于标题行列数，则后面几列都是空值
		}
		// 获取表内容数据
		Map<Integer, String> dataValueMap = getValueMap(dataList);
		ReturnExcelRowParseCommonDTO<?> parseResultDTO = BpoiExcelUtil.convertRowToObject(excelParseClass, maxColumn,
				rowIndex, dataValueMap, this.headMap, param.getErrInfoSj());
		if (!BpoiConstants.commonReturnStatus.SUCCESS.getValue().equals(parseResultDTO.getResultCode())) {
			throw new BpoiAlertException(param.getErrInfoSj().toString());
		}
		// 超出限定的错误提示计数，则停止解析，返回报错信息
		param.setErrHintCount(param.getErrHintCount() + parseResultDTO.getErrHintCount());
		if (param.getMaxErrHintCount() != null && param.getErrHintCount() >= param.getMaxErrHintCount()) {
			throw new BpoiAlertException(param.getErrInfoSj().toString());
		}
		// 获取解析出的数据
		Object dataT = parseResultDTO.getData();
		if (dataT != null && this.handler != null) {
			this.handler.handle(dataT);
		}
		if (this.handler == null) {
			this.objList.add(dataT);
		}

	}

	/**
	 * 初始化表头数据
	 * @param dataList
	 */
	private void initHeadData(List<ExcelSaxReadCellDTO> dataList) throws Exception {
		// 获取表头数据
		Map<Integer, String> headValueMap = getValueMap(dataList);
		// 验证并解析表头信息
		BpoiReturnCommonDTO<ExcelHeadValidateResultDTO> headValidateResult = BpoiExcelUtil.validateHead(
				headValueMap, this.excelParseClass, dataList.size());
		if (!BpoiConstants.commonReturnStatus.SUCCESS.getValue().equals(headValidateResult.getResultCode())) {
			throw new BpoiAlertException(headValidateResult.getErrMsg());
		}
		this.headMap = headValidateResult.getData().getHeadMap();
		this.maxColumn = headValidateResult.getData().getMaxColumn();
	}

	/**
	 * 获取该行数据对应的Map
	 * @param dataList 数据列表
	 * @return Map（Key：从0开始计的行号  Value：单元格内容）
	 */
	private Map<Integer, String> getValueMap(List<ExcelSaxReadCellDTO> dataList) {
		// 获取表内容数据
		Map<Integer, String> dataValueMap = new HashMap<>();
		for (int i = 0; i < dataList.size(); i++) {
			String headData = String.valueOf(dataList.get(i).getValue());
			if (headData != null && headData.length() > 0) {
				dataValueMap.put(i, headData);
			}
		}
		return dataValueMap;
	}
}
