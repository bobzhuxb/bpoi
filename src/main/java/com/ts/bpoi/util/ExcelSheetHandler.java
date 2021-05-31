package com.ts.bpoi.util;

import com.ts.bpoi.base.BpoiConstants;
import com.ts.bpoi.dto.ExcelSaxReadCellDTO;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Sheet的SAX解析处理器
 * XML格式请参见：https://docs.microsoft.com/zh-cn/office/open-xml/working-with-tables
 * @author Bob
 */
public class ExcelSheetHandler extends DefaultHandler {

	private SharedStringsTable sst;		// SST索引

	private String lastContents;		// 最新解析出的单元格内容

	private int curRow = 0;		// 当前行

	private int curCol = 0;		// 当前列

	private String lastCellId;	// 上个有内容的单元格id（例如A1、B1、A2等），用于判断空单元格

	private String lastRowId;	// 上一行id（例如1、2、3等）, 用于判断空行

	private boolean hasV = false;	// 判断单元格cell的c标签下是否有v，否则可能数据错位

	private BpoiConstants.excelCellValueType type;

	private IExcelSaxRowRead read;

	private List<ExcelSaxReadCellDTO> rowList = new ArrayList<>();	// 存储行记录的容器

	public ExcelSheetHandler(SharedStringsTable sst, IExcelSaxRowRead rowRead) {
		this.sst = sst;
		this.read = rowRead;
	}

	/**
	 * XML标签的起始回调
	 * @param uri
	 * @param localName
	 * @param name
	 * @param attributes
	 * @throws SAXException
	 */
	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// 置空
		lastContents = "";
		// 行开始
		if ("row".equals(name)) {
			String rowNum = attributes.getValue("r");
			// 判断前面或中间的空行（sheet尾部的空行不计）
			if (lastRowId != null) {
				// 与上一行相差超过1，说明中间有空行
				int gap = Integer.parseInt(rowNum) - Integer.parseInt(lastRowId);
				if (gap > 1) {
					gap -= 1;
					while (gap > 0) {
						gap--;
						// 行数加1（因为空行不会执行endElement方法）
						curRow++;
						// 处理空行（如果必要）
						read.parse(curRow, null);
					}
				}
			}
			// 记录保存当前行id
			lastRowId = attributes.getValue("r");
		}
		// c => 单元格
		if ("c".equals(name)) {
			String rowId = attributes.getValue("r");
			// 判断前面或中间的空单元格（行尾的空单元格不计）
			int nowColIndex = BpoiExcelUtil.getColumnIndexByCellName(rowId);
			// 注意，lastCellId在上一行结束的时候都会置空
			if (lastCellId != null) {
				// 该行非第一个有记录的单元格（index从1开始计）
				int lastColIndex = BpoiExcelUtil.getColumnIndexByCellName(lastCellId);
				// 计算同一行两个单元格之间的间隔
				int gap = nowColIndex - lastColIndex;
				for (int i = 0; i < gap - 1; i++) {
					// 间隔大于1，说明之间有空单元格，追加（gap - 1）条记录
					rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.String, null));
					curCol++;
				}
			} else {
				// 该行第一个有记录的单元格（index从1开始计），如果第一个单元格不在A列，则说明前面有空单元格
				for (int i = 0; i < nowColIndex - 1; i++) {
					// 追加（nowColIndex - 1）条记录
					rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.String, null));
					curCol++;
				}
			}
			// 记录当前已处理的单元格c标签
			lastCellId = rowId;

			// 如果下一个元素是 SST 的索引，则将nextIsString标记为true
			String cellType = attributes.getValue("t");
			if ("s".equals(cellType)) {
				type = BpoiConstants.excelCellValueType.String;
				return;
			}
			// 日期格式
			cellType = attributes.getValue("s");
			if ("1".equals(cellType)) {
				type = BpoiConstants.excelCellValueType.Date;
			} else if ("2".equals(cellType)) {
				type = BpoiConstants.excelCellValueType.Number;
			}
		} else if ("t".equals(name)) {
			// 当元素为t时
			type = BpoiConstants.excelCellValueType.TElement;
		}
	}

	/**
	 * XML标签的结束回调
	 * @param uri
	 * @param localName
	 * @param name
	 * @throws SAXException
	 */
	@Override
	public void endElement(String uri, String localName, String name) {
		// 根据SST的索引值的到单元格的真正要存储的字符串
		// 这时characters()方法可能会被调用多次
		if (BpoiConstants.excelCellValueType.String.equals(type)) {
			try {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			} catch (Exception e) {

			}
		}
		// 单元格结束，没有v时需要补位（有的单元格只修改单元格格式，而没有内容，会出现c标签下没有v标签）
		if ("c".equals(name)) {
			if (!hasV) {
				rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.String, null));
			}
			hasV = false;
		}
		if (BpoiConstants.excelCellValueType.TElement.equals(type)) {
			// t元素也包含字符串
			String value = lastContents.trim();
			rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.String, value));
			curCol++;
			type = BpoiConstants.excelCellValueType.None;
		} else if ("v".equals(name)) {
			// v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
			// 将单元格内容加入rowList中，在这之前先去掉字符串前后的空白符
			hasV = true;
			String value = lastContents.trim();
			value = value.equals("") ? " " : value;
			if (BpoiConstants.excelCellValueType.Date.equals(type)) {
				Date date = HSSFDateUtil.getJavaDate(Double.valueOf(value));
				rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.Date, date));
			} else if (BpoiConstants.excelCellValueType.Number.equals(type)) {
				BigDecimal bd = new BigDecimal(value);
				rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.Number, bd));
			} else if (BpoiConstants.excelCellValueType.String.equals(type)) {
				rowList.add(curCol, new ExcelSaxReadCellDTO(BpoiConstants.excelCellValueType.String, value));
			}
			curCol++;
		} else if ("row".equals(name)) {
			// 如果标签名称为row，这说明已到行尾，处理该行数据
			read.parse(curRow, rowList);
			// 清空改行数据
			rowList.clear();
			curRow++;
			// 初始化当前列号
			curCol = 0;
			// 上个有内容的单元格id（用于判断空单元格）
			lastCellId = null;
		}

	}

	/**
	 * XML标签的内容回调，该方法每个标签内解析可能会调用多次，拼装起来即可
	 * @param ch
	 * @param start
	 * @param length
	 * @throws SAXException
	 */
	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		// 得到单元格内容的值
		lastContents += new String(ch, start, length);
	}

}
