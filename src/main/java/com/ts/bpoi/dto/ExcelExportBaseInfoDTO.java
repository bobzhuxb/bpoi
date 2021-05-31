package com.ts.bpoi.dto;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.OutputStream;

public class ExcelExportBaseInfoDTO {

    private String fileName;            // 最终文件名

    private SXSSFWorkbook workbook;     // Excel工作簿

    private OutputStream outputStream;  // 输出流

    private String suffix;              // 后缀

    public ExcelExportBaseInfoDTO() {}

    public ExcelExportBaseInfoDTO(String fileName, SXSSFWorkbook workbook, OutputStream outputStream, String suffix) {
        this.fileName = fileName;
        this.workbook = workbook;
        this.outputStream = outputStream;
        this.suffix = suffix;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(SXSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }
}
