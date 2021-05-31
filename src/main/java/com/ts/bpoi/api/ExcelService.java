package com.ts.bpoi.api;

import com.ts.bpoi.dto.ExcelExportDTO;
import com.ts.bpoi.dto.BpoiReturnCommonDTO;
import com.ts.bpoi.util.ExcelServiceTemplateUtil;
import com.ts.bpoi.util.IExcelReadRowHandler;
import com.ts.bpoi.util.IExcelSaxRowRead;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.List;

/**
 * 共通方法
 * @author Bob
 */
public class ExcelService {

    /**
     * 为了打包成功而加，没有任何处理
     * @param args
     */
    public static void main(String[] args) {}

    /**
     * 导出单Sheet的Excel
     * @param response HTTP Response
     * @param excelExportDTO 导出的Excel相关数据
     * @return 导出结果
     * @throws Exception
     */
    public static void exportExcel(HttpServletResponse response, ExcelExportDTO excelExportDTO) {
        new ExcelServiceTemplateUtil().exportExcel(response, excelExportDTO);
    }

    /**
     * 导出多Sheet的Excel
     * @param response HTTP Response
     * @param excelExportDTOList 导出的Excel相关数据（每个元素表示一个Sheet的描述）
     * @return 导出结果
     * @throws Exception
     */
    public static void exportExcel(HttpServletResponse response, List<ExcelExportDTO> excelExportDTOList) {
        new ExcelServiceTemplateUtil().exportExcel(response, excelExportDTOList);
    }

    /**
     * 对Excel进行加密（读写均加密）
     * @param fullFileName 带全路径的文件名
     * @param password 密码
     * @throws Exception
     */
    public static void encryptExcel(String fullFileName, String password) throws Exception {
        new ExcelServiceTemplateUtil().encryptExcel(fullFileName, password);
    }

    /**
     * 通用解析Excel文件（自主选择使用大文件解析方式还是小文件解析方式）
     * @param fullFileName 本地全路径的文件名
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public static <T> BpoiReturnCommonDTO<List<T>> importParseExcel(String fullFileName, Class<T> excelParseClass,
                                                             Integer maxErrHintCount) {
        return new ExcelServiceTemplateUtil().importParseExcel(fullFileName, excelParseClass, maxErrHintCount);
    }

    /**
     * 通用解析Excel文件（自主选择使用大文件解析方式还是小文件解析方式）
     * @param excelType Excel类型（xls或xlsx）
     * @param fileInputStream 文件输入流
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public static <T> BpoiReturnCommonDTO<List<T>> importParseExcel(String excelType, InputStream fileInputStream,
                                                             Class<T> excelParseClass, Integer maxErrHintCount) {
        return new ExcelServiceTemplateUtil().importParseExcel(excelType, fileInputStream, excelParseClass,
                maxErrHintCount);
    }

    /**
     * 通用解析Excel文件（小文件）
     * @param fullFileName 本地全路径的文件名
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public static <T> BpoiReturnCommonDTO<List<T>> importParseExcelSmall(String fullFileName, Class<T> excelParseClass,
                                                                  Integer maxErrHintCount) {
        return new ExcelServiceTemplateUtil().importParseExcelSmall(fullFileName, excelParseClass,
                maxErrHintCount);
    }

    /**
     * 通用解析Excel文件（小文件）
     * @param excelType Excel类型（xls或xlsx）
     * @param fileInputStream 文件输入流
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public static <T> BpoiReturnCommonDTO<List<T>> importParseExcelSmall(String excelType, InputStream fileInputStream,
                                                                  Class<T> excelParseClass, Integer maxErrHintCount) {
        return new ExcelServiceTemplateUtil().importParseExcelSmall(excelType, fileInputStream,
                excelParseClass, maxErrHintCount);
    }

    /**
     * 通用解析Excel文件（大文件）
     * @param fullFileName 本地全路径的文件名
     * @param excelParseClass 解析用Class
     * @param rowRead SAX解析读取到的行
     * @param handler 自定义行处理器
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public static <T> BpoiReturnCommonDTO<List<T>> importParseExcelLarge(String fullFileName, Class<T> excelParseClass,
                                                                  IExcelSaxRowRead rowRead, IExcelReadRowHandler handler,
                                                                  Integer maxErrHintCount) {
        return new ExcelServiceTemplateUtil().importParseExcelLarge(fullFileName, excelParseClass,
                rowRead, handler, maxErrHintCount);
    }

    /**
     * 通用解析Excel文件（大文件）
     * @param excelType Excel类型（xls或xlsx）
     * @param fileInputStream 文件输入流
     * @param excelParseClass 解析用Class
     * @param rowRead SAX解析读取到的行
     * @param handler 自定义行处理器
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public static <T> BpoiReturnCommonDTO<List<T>> importParseExcelLarge(String excelType, InputStream fileInputStream,
                                                                  Class<T> excelParseClass, IExcelSaxRowRead rowRead,
                                                                  IExcelReadRowHandler handler, Integer maxErrHintCount) {
        return new ExcelServiceTemplateUtil().importParseExcelLarge(excelType, fileInputStream,
                excelParseClass, rowRead, handler, maxErrHintCount);
    }

}
