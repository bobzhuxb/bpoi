package com.ts.bpoi.base;

import org.apache.poi.ss.usermodel.IndexedColors;

public class BpoiConstants {

    /**
     * Excel导出中的主题色
     */
    public static final short EXCEL_THEME_COLOR = IndexedColors.GREY_25_PERCENT.getIndex();

    /**
     * 大Excel的指定字节数（只针对xlsx文件）
     */
    public static final long EXCEL_LARGE_BYTES = 1024 * 1024;

    /**
     * 通用Enum接口
     */
    public interface EnumInter {
        Object getValue();
        String getName();
    }

    /**
     * 系统通用返回状态
     */
    public enum commonReturnStatus implements EnumInter {
        SUCCESS("1", "操作成功"), FAIL("2", "操作失败");
        private String value;
        private String name;
        private commonReturnStatus(String value, String name) {
            this.value = value;
            this.name = name;
        }
        @Override
        public String getValue(){
            return value;
        }
        @Override
        public String getName(){
            return name;
        }
    }

    /**
     * Excel的类型
     */
    public enum excelType implements EnumInter {
        XLS("xls", "2003版本"), XLSX("xlsx", "2007版本");
        private String value;
        private String name;
        private excelType(String value, String name) {
            this.value = value;
            this.name = name;
        }
        @Override
        public String getValue(){
            return value;
        }
        @Override
        public String getName(){
            return name;
        }
    }

    /**
     * Excel单元格的值类型
     */
    public enum excelCellValueType {
        String, Number, Boolean, Date, TElement, Null, None;
    }

}
