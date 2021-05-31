package com.ts.bpoi.util;

/**
 * 解析出的Excel行处理接口
 * @author Bob
 */
public interface IExcelReadRowHandler<T> {

	/**
	 * 处理解析出的行对象
	 * @param data 解析出的行对象
	 */
	void handle(T data);

}
