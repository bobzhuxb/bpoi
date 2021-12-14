package com.ts.bpoi.util;

import com.ts.bpoi.base.BpoiConstants;
import com.ts.bpoi.dto.*;
import com.ts.bpoi.error.BpoiAlertException;
import com.ts.bpoi.error.BpoiException;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

/**
 * 共通方法
 * @author Bob
 */
public class ExcelServiceTemplateUtil {

    private final Logger log = LoggerFactory.getLogger(ExcelServiceTemplateUtil.class);

    /**
     * 导出单Sheet的Excel
     * 注意：1、Spring框架中，不要对HttpServletResponse对象的ServletOutputStream做流关闭处理
     * 2、对于下载文件的方法请不要有任何返回值，因为我们写完业务逻辑后Spring框架层还会做一些额外的工作（可能会用到ServletOutputStream对象）
     * @param response HTTP Response
     * @param excelExportDTO 导出的Excel相关数据
     * @return 导出结果
     * @throws Exception
     */
    public void exportExcel(HttpServletResponse response, ExcelExportDTO excelExportDTO) {
        SXSSFWorkbook workbook = null;
        OutputStream outputStream = null;
        try {
            // 生成Excel信息
            ExcelExportBaseInfoDTO excelExportBaseInfoDTO = formExcelForExport(response, excelExportDTO);
            // 获取解析参数
            workbook = excelExportBaseInfoDTO.getWorkbook();
            outputStream = excelExportBaseInfoDTO.getOutputStream();
            String fileName = excelExportBaseInfoDTO.getFileName();
            // 生成Sheet信息
            // sheet名
            String sheetName = excelExportDTO.getSheetName();
            if (sheetName == null || "".equals(sheetName)) {
                sheetName = fileName;
            }
            formExcelSheetForExport(excelExportDTO, sheetName, workbook);
            // 写入数据
            workbook = doWriteExcelStream(workbook, outputStream, excelExportBaseInfoDTO.getSuffix(),
                    excelExportDTO.getOpenPassword());
        } catch (Exception e) {
            throw BpoiException.errWithDetail("导出失败", e.getMessage(), e);
        } finally {
            // 关闭
            try {
                if (response == null) {
                    // 纯写文件需要关闭流，导出时不需要
                    if (workbook != null) {
                        workbook.close();
                    }
                    if (outputStream != null) {
                        outputStream.close();
                    }
                }
            } catch (Exception e) {
                log.warn("Excel导出，流关闭失败", e.getMessage());
            }
        }
    }

    /**
     * 导出多Sheet的Excel
     * 注意：1、Spring框架中，不要对HttpServletResponse对象的ServletOutputStream做流关闭处理
     * 2、对于下载文件的方法请不要有任何返回值，因为我们写完业务逻辑后Spring框架层还会做一些额外的工作（可能会用到ServletOutputStream对象）
     * @param response HTTP Response
     * @param excelExportDTOList 导出的Excel相关数据（每个元素表示一个Sheet的描述）
     * @return 导出结果
     * @throws Exception
     */
    public void exportExcel(HttpServletResponse response, List<ExcelExportDTO> excelExportDTOList) {
        // 验证参数
        if (excelExportDTOList == null && excelExportDTOList.size() == 0) {
            throw new BpoiAlertException("Excel多sheet导出参数异常");
        }
        if (excelExportDTOList.size() == 1) {
            // 单sheet的Excel转入对应的方法进行处理
            exportExcel(response, excelExportDTOList.get(0));
            return;
        }
        SXSSFWorkbook workbook = null;
        OutputStream outputStream = null;
        try {
            // 与Excel文件本身相关的属性（非sheet的属性）写在第一个元素中
            ExcelExportBaseInfoDTO excelExportBaseInfoDTO = formExcelForExport(response, excelExportDTOList.get(0));
            // 获取解析参数
            workbook = excelExportBaseInfoDTO.getWorkbook();
            outputStream = excelExportBaseInfoDTO.getOutputStream();
            // 生成每一个Sheet信息
            for (ExcelExportDTO excelExportDTO : excelExportDTOList) {
                // sheet名
                String sheetName = excelExportDTO.getSheetName();
                if (sheetName == null || "".equals(sheetName)) {
                    throw new BpoiAlertException("参数错误，部分sheet名未配置");
                }
                formExcelSheetForExport(excelExportDTO, sheetName, workbook);
            }
            // 写入数据
            workbook = doWriteExcelStream(workbook, outputStream, excelExportBaseInfoDTO.getSuffix(),
                    excelExportDTOList.get(0).getOpenPassword());
        } catch (BpoiAlertException e) {
            throw e;
        } catch (Exception e) {
            throw BpoiException.errWithDetail("导出失败", e.getMessage(), e);
        } finally {
            // 关闭
            try {
                if (response == null) {
                    // 纯写文件需要关闭流，导出时不需要
                    if (workbook != null) {
                        workbook.close();
                    }
                    if (outputStream != null) {
                        outputStream.close();
                    }
                }
            } catch (Exception e) {
                log.warn("Excel导出，流关闭失败", e.getMessage());
            }
        }
    }

    /**
     * 为Excel导出解析Excel的基本信息
     * @param response HTTP Response
     * @param excelExportDTO 导出的Excel相关数据
     * @return Excel的基本信息
     * @throws Exception
     */
    private ExcelExportBaseInfoDTO formExcelForExport(HttpServletResponse response, ExcelExportDTO excelExportDTO) throws Exception {
        // Excel文件名（包含相对路径，但不包含后缀）
        // 如果response不为空，则fileName没有相对路径；如果response为空，则fileName有相对路径
        // response不为空表示导出文件到客户端，response为空表示导出文件到服务器本地
        // 最终文件名
        String fileName = null;
        // 相对路径
        String relativePath = null;
        // 包含相对路径，但不包含后缀的文件名
        String relativeFileName = excelExportDTO.getFileName();
        if (response != null) {
            fileName = relativeFileName;
        } else {
            fileName = relativeFileName.substring(relativeFileName.lastIndexOf(File.separator) + 1);
            relativePath = relativeFileName.substring(0, relativeFileName.lastIndexOf(File.separator));
        }
        // 后缀名
        String suffix = BpoiConstants.excelType.XLSX.getValue();
        if (excelExportDTO.getExcelType() != null) {
            suffix = excelExportDTO.getExcelType();
        }
        // 创建Excel工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        if (response != null) {
            // 设置头信息
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");
            // 文件名进行编码
            response.setHeader("Content-Disposition", "attachment;filename*=utf-8'zh_cn'"
                    + URLEncoder.encode(fileName.replace("+", " ") + "." + suffix, "UTF-8"));
        }
        // 输出流
        OutputStream outputStream = null;
        if (response != null) {
            // 创建一个输出流
            outputStream = response.getOutputStream();
        } else {
            // excelExportDTO.getFileName()是相对路径
            String baseLocation = excelExportDTO.getBaseLocation();
            if (baseLocation == null || "".equals(baseLocation.trim())) {
                // 默认输出路径
                baseLocation = "." + File.separator;
            }
            if (!baseLocation.endsWith(File.separator)) {
                baseLocation += File.separator;
            }
            outputStream = new FileOutputStream(baseLocation + relativePath
                    + File.separator + fileName + "." + suffix);
        }
        return new ExcelExportBaseInfoDTO(fileName, workbook, outputStream, suffix);
    }

    /**
     * 为Excel导出解析Excel中某个sheet的基本信息
     * @param excelExportDTO 导出的Excel相关数据
     * @param sheetName sheet名
     * @param workbook Excel工作簿
     * @throws Exception
     */
    private void formExcelSheetForExport(ExcelExportDTO excelExportDTO, String sheetName, SXSSFWorkbook workbook) throws Exception {
        // 最大列数（用于自适应列宽）
        int maxColumn = excelExportDTO.getMaxColumn();
        // 实际表格（包括标题行）的开始行
        int tableStartRow = excelExportDTO.getTableStartRow();
        // 在实际表格前面部分的单元格
        List<ExcelCellDTO> beforeDataCellList = excelExportDTO.getBeforeDataCellList();
        // 在实际表格后面部分的单元格
        List<ExcelCellDTO> afterDataCellList = excelExportDTO.getAfterDataCellList();
        // 标题行
        List<ExcelTitleDTO> titleList = excelExportDTO.getTitleList();
        // 数据
        List<?> dataList = excelExportDTO.getDataList();
        // 要合并的单元格
        List<ExcelCellRangeDTO> cellRangeList = excelExportDTO.getCellRangeList();

        // ==================== 开始生成Excel ==========================
        // 创建sheet页
        SXSSFSheet sheet = null;
        try {
            sheet = workbook.createSheet(sheetName);
        } catch (Exception e) {
            // 出现了重名现象，则创建非指定的名称Sheet
            sheet = workbook.createSheet();
        }
        // 如果设置了工作表保护，则设定密码
        if (excelExportDTO.getSheetProtectPassword() != null && !"".equals(excelExportDTO.getSheetProtectPassword().trim())) {
            sheet.protectSheet(excelExportDTO.getSheetProtectPassword());
        }
        // 存储最大列宽，用于处理中文不能自动调整列宽的问题
        Map<Integer, Integer> maxWidthMap = new HashMap<>();
        // 预处理合并单元格（以防出现Attempting to write ... that is already written to disk.）
        // Key：要合并的单元格的最后一行  Value：最后一行是该数值的所有合并单元格
        Map<Integer, List<ExcelCellRangeDTO>> lastRowToMergeMap = BpoiExcelUtil.mergeCellGroup(cellRangeList);
        // 在实际表格上方的单元格
        if (beforeDataCellList != null) {
            // key：relativeRow相对行数，value：该行的单元格数据
            BpoiExcelUtil.addDataToExcel(beforeDataCellList, 0, maxWidthMap, lastRowToMergeMap,
                    cellRangeList, excelExportDTO.getWrapSpecial(), workbook, sheet);
        }
        if (titleList != null && titleList.size() > 0) {
            // 设定了titleList
            // 如果title没有设定列序号，则从0开始按顺序设置序列号
            if (titleList.get(0).getColIndex() == null) {
                for (int i = 0; i < titleList.size(); i++) {
                    titleList.get(i).setColIndex(i);
                }
            }
        } else {
            // 没有设定titleList
            // 如果有标题列，则根据数据列类型对应的注解反推titleList（只有非Map类型且字段有@ExcelProperty注解的有效）
            if (dataList != null && dataList.size() > 0) {
                Object dataObj0 = dataList.get(0);
                if (dataObj0 instanceof Map) {
                    // do nothing
                } else {
                    // 根据数据类型的字段注解@ExcelProperty反推titleList
                    titleList = BpoiExcelUtil.getTitleListFromObj(dataObj0.getClass());
                    excelExportDTO.setTitleList(titleList);
                }
            }
        }
        if (tableStartRow >= 0 && titleList != null && titleList.size() > 0
                && dataList != null && dataList.size() > 0) {
            // 存在有数据列表的情况下，生成数据列表
            // 创建表头
            SXSSFRow headRow = sheet.createRow(tableStartRow);
            // 表头统一居中，背景为灰色，加边框，可换行
            CellStyle titleNameStyle = workbook.createCellStyle();
            BpoiExcelUtil.setAlignment(titleNameStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
            BpoiExcelUtil.setBackgroundColor(titleNameStyle, BpoiConstants.EXCEL_THEME_COLOR);
            BpoiExcelUtil.setBorder(titleNameStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
            // 可自动换行
            BpoiExcelUtil.setWrapText(titleNameStyle, true);
            // 设置表头信息（Key：标题的字段名  Value：从0开始计的列号）
            Map<String, Integer> titleNameMap = new HashMap<>();
            for (int index = 0; index < titleList.size(); index++) {
                // index是数组中的序号
                ExcelTitleDTO titleDTO = titleList.get(index);
                // 获取列号
                int column = titleDTO.getColIndex();
                titleNameMap.put(titleDTO.getTitleName(), column);
                SXSSFCell cell = headRow.createCell(column);
                // 设置表头内容（替换自动换行标识符）
                String titleContent = titleDTO.getTitleContent() == null ? "" : titleDTO.getTitleContent();
                if (excelExportDTO.getWrapSpecial() != null) {
                    cell.setCellValue(titleContent.replace(excelExportDTO.getWrapSpecial(), "\n"));
                }
                cell.setCellStyle(titleNameStyle);
                // 每个单元格都要计算该列的最大宽度
                BpoiExcelUtil.computeMaxColumnWith(maxWidthMap, cell, tableStartRow, column, null, cellRangeList);
            }
            // 默认表数据统一居中，加边框，可换行
            CellStyle defaultDataStyle = workbook.createCellStyle();
            BpoiExcelUtil.setAlignment(defaultDataStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
            BpoiExcelUtil.setBorder(defaultDataStyle, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN, BorderStyle.THIN);
            // 可自动换行
            BpoiExcelUtil.setWrapText(defaultDataStyle, true);
            // 判断是否有合并单元格，有的话就合并（此时所在行为tableStartRow）
            List<ExcelCellRangeDTO> tableStartMergeList = lastRowToMergeMap.get(tableStartRow);
            if (tableStartMergeList != null && tableStartMergeList.size() > 0) {
                for (ExcelCellRangeDTO cellRangeDTO : tableStartMergeList) {
                    // 合并单元格
                    BpoiExcelUtil.doMergeCell(sheet, cellRangeDTO);
                }
            }
            // 特殊格式的单元格List
            List<ExcelCellDTO> specialStyleCellList = excelExportDTO.getSpecialStyleCellList();
            // 特殊格式的单元格Map（Key：相对行序号，从数据的第一行开始记为0  Value：特殊格式的单元格列表）
            Map<Integer, List<ExcelCellDTO>> specialStyleCellMap = new HashMap<>();
            if (specialStyleCellList != null && specialStyleCellList.size() > 0) {
                for (ExcelCellDTO specialStyleCell : specialStyleCellList) {
                    int relativeRow = specialStyleCell.getRelativeRow();
                    List<ExcelCellDTO> rowSpecialStyleCellList = specialStyleCellMap.get(relativeRow);
                    if (rowSpecialStyleCellList == null) {
                        rowSpecialStyleCellList = new ArrayList<>();
                        specialStyleCellMap.put(relativeRow, rowSpecialStyleCellList);
                    }
                    rowSpecialStyleCellList.add(specialStyleCell);
                }
            }
            // 填入表数据
            for (int dataRowCount = 0; dataRowCount < dataList.size(); dataRowCount++) {
                // 当前行
                int nowRow = tableStartRow + dataRowCount + 1;
                SXSSFRow dataRow = sheet.createRow(nowRow);
                Object dataObject = dataList.get(dataRowCount);
                // 特殊格式的单元格列表
                List<ExcelCellDTO> rowSpecialStyleCellList = specialStyleCellMap.get(dataRowCount);
                if (dataObject instanceof Map) {
                    // Map获取数据
                    for (int index = 0; index < titleList.size(); index++) {
                        // 获取字段名
                        String titleName = titleList.get(index).getTitleName();
                        // 获取列号（从0开始计）
                        int column = titleNameMap.get(titleName);
                        Object data = ((Map) dataObject).get(titleName);
                        // 设置数据列的数据和格式
                        setDataAndStyle(excelExportDTO, data, dataRow, nowRow, column, defaultDataStyle,
                                rowSpecialStyleCellList, maxWidthMap, cellRangeList, workbook);
                    }
                } else {
                    // 反射获取对象属性
                    for (int index = 0; index < titleList.size(); index++) {
                        // 获取字段名
                        String titleName = titleList.get(index).getTitleName();
                        // 获取列号（从0开始计）
                        int column = titleNameMap.get(titleName);
                        // 获取属性（使用apache的包可以获取包括父类的属性）
                        Field field = FieldUtils.getField(dataObject.getClass(), titleName, true);
                        // 设置对象的访问权限，保证对private的属性的访问
                        if (field != null) {
                            field.setAccessible(true);
                            Object data = field.get(dataObject);
                            // 设置数据列的数据和格式
                            setDataAndStyle(excelExportDTO, data, dataRow, nowRow, column, defaultDataStyle,
                                    rowSpecialStyleCellList, maxWidthMap, cellRangeList, workbook);
                        }
                    }
                }
                // 每处理一行都要判断是否有合并单元格，有的话就合并（此时所在行为nowRow）
                List<ExcelCellRangeDTO> dataRowMergeList = lastRowToMergeMap.get(nowRow);
                if (dataRowMergeList != null && dataRowMergeList.size() > 0) {
                    for (ExcelCellRangeDTO cellRangeDTO : dataRowMergeList) {
                        // 合并单元格
                        BpoiExcelUtil.doMergeCell(sheet, cellRangeDTO);
                    }
                }
            }
        }
        // 在实际表格下方的单元格
        if (afterDataCellList != null) {
            // key：relativeRow相对行数，value：该行的单元格数据
            BpoiExcelUtil.addDataToExcel(afterDataCellList, tableStartRow + dataList.size() + 1, maxWidthMap,
                    lastRowToMergeMap, cellRangeList, excelExportDTO.getWrapSpecial(), workbook, sheet);
        }
        // 设置为根据内容自动调整列宽，必须在单元格设值以后进行
        sheet.trackAllColumnsForAutoSizing();
        // 如果没有设定最大列，则根据实际数据计算
        if (maxColumn == 0) {
            maxColumn = BpoiExcelUtil.computeMaxColumnIndex(excelExportDTO);
        }
        for (int column = 0; column < maxColumn; column++) {
            sheet.autoSizeColumn(column);
            // 处理中文不能自动调整列宽的问题
            if (maxWidthMap.get(column) != null) {
                sheet.setColumnWidth(column, maxWidthMap.get(column));
            }
        }
    }

    /**
     * 设置数据列的数据和格式
     * @param excelExportDTO 导出的Excel相关数据
     * @param data 数据值
     * @param dataRow 数据行
     * @param nowRow 当前行
     * @param column 当前列
     * @param defaultDataStyle 默认数据格式
     * @param rowSpecialStyleCellList 当前行特殊格式的单元格列表
     * @param maxWidthMap 存储的最大列宽Map
     * @param cellRangeList 要合并的单元格
     * @param workbook Excel工作簿
     */
    private static void setDataAndStyle(ExcelExportDTO excelExportDTO, Object data, SXSSFRow dataRow, int nowRow,
                                        int column, CellStyle defaultDataStyle, List<ExcelCellDTO> rowSpecialStyleCellList,
                                        Map<Integer, Integer> maxWidthMap, List<ExcelCellRangeDTO> cellRangeList,
                                        SXSSFWorkbook workbook) {
        SXSSFCell cell = dataRow.createCell(column);
        // 设置数据（替换自动换行标识符）
        if (excelExportDTO.getWrapSpecial() != null) {
            cell.setCellValue(data == null ? "" : data.toString().replace(excelExportDTO.getWrapSpecial(), "\n"));
        }
        // 指定列格式
        ExcelCellDTO dataCellStyleFromColumn = excelExportDTO.getDataSpecialStyleMap() == null ? null
                : excelExportDTO.getDataSpecialStyleMap().get(column);
        CellStyle dataStyle = defaultDataStyle;
        // 如果指定过某列的格式，则用该格式
        if (dataCellStyleFromColumn != null) {
            dataStyle = BpoiExcelUtil.setCellStyle(dataCellStyleFromColumn, workbook);
        }
        // 如果指定过具体单元格格式，则用该格式
        if (rowSpecialStyleCellList != null && rowSpecialStyleCellList.size() > 0) {
            for (ExcelCellDTO rowSpecialStyleCell : rowSpecialStyleCellList) {
                if (rowSpecialStyleCell.getColumn() == column) {
                    dataStyle = BpoiExcelUtil.setCellStyle(rowSpecialStyleCell, workbook);
                    break;
                }
            }
        }
        cell.setCellStyle(dataStyle);
        // 每个单元格都要计算该列的最大宽度
        BpoiExcelUtil.computeMaxColumnWith(maxWidthMap, cell, nowRow, column, null, cellRangeList);
    }

    /**
     * 将workbook写入流，如果需要加密，则加密后写入流
     * @param workbook Excel工作簿
     * @param outputStream 输出流
     * @param suffix 后缀
     * @param encryptPassword 加密密码（若为null，则表示加密）
     */
    private SXSSFWorkbook doWriteExcelStream(SXSSFWorkbook workbook, OutputStream outputStream, String suffix,
                                             String encryptPassword) throws Exception {
        if (encryptPassword == null || "".equals(encryptPassword.trim())) {
            // 不加密，直接写入流
            workbook.write(outputStream);
            return workbook;
        }
        // 先把流写入临时文件，然后加密
        String tempFileName = System.currentTimeMillis() + "_" + UUID.randomUUID().toString() + "." + suffix;
        OutputStream tempOutputStream = null;
        try {
            tempOutputStream = new FileOutputStream(tempFileName);
            workbook.write(tempOutputStream);
        } finally {
            if (tempOutputStream != null) {
                tempOutputStream.close();
            }
        }
        encryptExcel(tempFileName, encryptPassword);
        // 将加密后的文件写回流
        InputStream encryptFileInputStream = null;
        try {
            encryptFileInputStream = new FileInputStream(tempFileName);
            byte[] read = new byte[10240];
            int length;
            // 开始读取
            while ((length = encryptFileInputStream.read(read)) != -1) {
                // 开始写入
                outputStream.write(read, 0, length);
            }
        } finally {
            if (null != encryptFileInputStream) {
                encryptFileInputStream.close();
            }
            // 删除临时的加密文件
            new File(tempFileName).delete();
        }
        return workbook;
    }

    /**
     * 对Excel进行加密（读写均加密）
     * @param fullFileName 带全路径的文件名
     * @param password 密码
     * @throws Exception
     */
    public void encryptExcel(String fullFileName, String password) throws Exception {
        POIFSFileSystem fs = new POIFSFileSystem();
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
        // EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);
        Encryptor enc = info.getEncryptor();
        // 设置密码
        enc.confirmPassword(password);
        OPCPackage opc = null;
        try {
            opc = OPCPackage.open(new File(fullFileName), PackageAccess.READ_WRITE);
            OutputStream os = enc.getDataStream(fs);
            opc.save(os);
        } finally {
            if (opc != null) {
                opc.close();
            }
        }
        // 保存加密后的文件
        try (FileOutputStream fos = new FileOutputStream(fullFileName)) {
            fs.writeFilesystem(fos);
        }
    }

    /**
     * 通用解析Excel文件
     * @param fullFileName 本地全路径的文件名
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public <T> BpoiReturnCommonDTO<List<T>> importParseExcel(String fullFileName,
                                                             Class<T> excelParseClass, Integer maxErrHintCount) {
        if (fullFileName == null || "".equals(fullFileName)) {
            return new BpoiReturnCommonDTO(BpoiConstants.commonReturnStatus.FAIL.getValue(), "导入文件名为空");
        }
        try {
            InputStream fileInputStream = new FileInputStream(fullFileName);
            String excelType = fullFileName.substring(fullFileName.lastIndexOf(".") + 1);
            return importParseExcel(excelType, fileInputStream, excelParseClass, maxErrHintCount);
        } catch (IOException e) {
            return new BpoiReturnCommonDTO(BpoiConstants.commonReturnStatus.FAIL.getValue(), "文件读取失败");
        }
    }

    /**
     * 通用解析Excel文件
     * @param excelType Excel类型（xls或xlsx）
     * @param fileInputStream 文件输入流
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public <T> BpoiReturnCommonDTO<List<T>> importParseExcel(String excelType, InputStream fileInputStream,
                                                             Class<T> excelParseClass, Integer maxErrHintCount) {
        if (BpoiConstants.excelType.XLS.getValue().equals(excelType)) {
            // 2003版本的解析只支持一种方案
            return importParseExcelSmall(excelType, fileInputStream, excelParseClass, maxErrHintCount);
        } else {
            // 2007版本的解析支持两种方案，自动选择
            try {
                if (fileInputStream.available() < BpoiConstants.EXCEL_LARGE_BYTES) {
                    // 小文件解析
                    return importParseExcelSmall(excelType, fileInputStream, excelParseClass, maxErrHintCount);
                } else {
                    // 大文件解析（默认不逐行处理）
                    return importParseExcelLarge(excelType, fileInputStream, excelParseClass, null, null, maxErrHintCount);
                }
            } catch (IOException e) {
                return new BpoiReturnCommonDTO(BpoiConstants.commonReturnStatus.FAIL.getValue(), "文件读取失败");
            }
        }
    }

    /**
     * 通用解析Excel文件（小文件）
     * @param fullFileName 本地全路径的文件名
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public <T> BpoiReturnCommonDTO<List<T>> importParseExcelSmall(String fullFileName, Class<T> excelParseClass,
                                                                  Integer maxErrHintCount) {
        if (fullFileName == null || "".equals(fullFileName)) {
            return new BpoiReturnCommonDTO(BpoiConstants.commonReturnStatus.FAIL.getValue(), "导入文件名为空");
        }
        try {
            InputStream fileInputStream = new FileInputStream(fullFileName);
            String excelType = fullFileName.substring(fullFileName.lastIndexOf(".") + 1);
            return importParseExcelSmall(excelType, fileInputStream, excelParseClass, maxErrHintCount);
        } catch (IOException e) {
            return new BpoiReturnCommonDTO(BpoiConstants.commonReturnStatus.FAIL.getValue(), "文件读取失败");
        }
    }

    /**
     * 通用解析Excel文件（小文件）
     * @param excelType Excel类型（xls或xlsx）
     * @param fileInputStream 文件输入流
     * @param excelParseClass 解析用Class
     * @param maxErrHintCount 最大错误提示数
     * @return 解析后的数据列表
     */
    public <T> BpoiReturnCommonDTO<List<T>> importParseExcelSmall(String excelType, InputStream fileInputStream,
                                                                  Class<T> excelParseClass, Integer maxErrHintCount) {
        try {
            if (!BpoiConstants.excelType.XLS.getValue().equals(excelType)
                    && !BpoiConstants.excelType.XLSX.getValue().equals(excelType)) {
                throw new BpoiAlertException("Excel文件格式错误，后缀必须为.xls或.xlsx");
            }
            Workbook workbook = null;
            Sheet sheet = null;
            Row firstRow = null;
            if (BpoiConstants.excelType.XLS.getValue().equals(excelType)) {
                workbook = new HSSFWorkbook(fileInputStream);
                sheet = workbook.getSheetAt(0);
                // 第一行（列名）
                firstRow = sheet.getRow(0);
            } else {
                workbook = new XSSFWorkbook(fileInputStream);
                sheet = workbook.getSheetAt(0);
                // 第一行（列名）
                firstRow = sheet.getRow(0);
            }
            // 表头最大列数
            int headMaxColumn = firstRow.getLastCellNum();
            // 表头数据Map（Key：从0开始计的列号 Value：列内容）
            Map<Integer, String> headValueMap = new HashMap<>();
            for (int col = 0; col < headMaxColumn; col++) {
                String cellValue = BpoiExcelUtil.getCellValueOfExcel(firstRow.getCell(col), null);
                headValueMap.put(col, cellValue);
            }
            // 验证并解析表头信息
            BpoiReturnCommonDTO<ExcelHeadValidateResultDTO> headValidateResult = BpoiExcelUtil.validateHead(
                    headValueMap, excelParseClass, headMaxColumn);
            if (!BpoiConstants.commonReturnStatus.SUCCESS.getValue().equals(headValidateResult.getResultCode())) {
                return new BpoiReturnCommonDTO<>(headValidateResult.getResultCode(), headValidateResult.getErrMsg());
            }
            Map<Integer, ExcelPropertyDataDTO> excelPropertyDataMap = headValidateResult.getData().getHeadMap();
            int maxColumn = headValidateResult.getData().getMaxColumn();
            // 统计错误提示数，不超过最大提示数
            int errHintCount = 0;
            StringJoiner errInfoSj = new StringJoiner("、");
            // 解析出的数据行内容
            List<T> dataList = new ArrayList<>();
            // 读取数据行的每一行内容进行解析
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                // 获取这一行
                Row dataRow = sheet.getRow(i);
                int dataMaxCell = dataRow.getLastCellNum();
                if (dataMaxCell > maxColumn) {
                    errInfoSj.add("第" + (i + 1) + "行列数超过" + maxColumn);
                    errHintCount++;
                    if (maxErrHintCount != null && errHintCount >= maxErrHintCount) {
                        return new BpoiReturnCommonDTO<>(BpoiConstants.commonReturnStatus.FAIL.getValue(), errInfoSj.toString());
                    }
                    continue;
                } else if (dataMaxCell < maxColumn) {
                    // 数据行列数小于标题行列数，则后面几列都是空值
                }
                // 当前行数据Map（Key：从0开始计的列号 Value：列内容）
                Map<Integer, String> dataValueMap = new HashMap<>();
                for (int col = 1; col < maxColumn + 1; col++) {
                    // 属性数据
                    ExcelPropertyDataDTO excelPropertyDataDTO = excelPropertyDataMap.get(col);
                    String cellValue = "";
                    if (col <= dataMaxCell) {
                        // 获取单元格中填入的值
                        cellValue = BpoiExcelUtil.getCellValueOfExcel(dataRow.getCell(col - 1), excelPropertyDataDTO.getDateFormat());
                    }
                    dataValueMap.put(col - 1, cellValue);
                }
                // 解析当前行的数据
                ReturnExcelRowParseCommonDTO<T> parseResultDTO = BpoiExcelUtil.convertRowToObject(excelParseClass, maxColumn,
                        i, dataValueMap, excelPropertyDataMap, errInfoSj);
                if (!BpoiConstants.commonReturnStatus.SUCCESS.getValue().equals(parseResultDTO.getResultCode())) {
                    return new BpoiReturnCommonDTO<>(BpoiConstants.commonReturnStatus.FAIL.getValue(), errInfoSj.toString());
                }
                // 超出限定的错误提示计数，则停止解析，返回报错信息
                errHintCount += parseResultDTO.getErrHintCount();
                if (maxErrHintCount != null && errHintCount >= maxErrHintCount) {
                    return new BpoiReturnCommonDTO<>(BpoiConstants.commonReturnStatus.FAIL.getValue(), errInfoSj.toString());
                }
                // 将解析出的这一行的数据添加到列表中
                dataList.add(parseResultDTO.getData());
            }
            // 错误信息整体返回
            if (errInfoSj.length() > 0) {
                return new BpoiReturnCommonDTO<>(BpoiConstants.commonReturnStatus.FAIL.getValue(), errInfoSj.toString());
            }
            // 返回解析后的全部数据
            return new BpoiReturnCommonDTO(dataList);
        } catch (Exception e) {
            throw BpoiException.errWithDetail("导入文件解析失败", e.getMessage(), e);
        }
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
    public <T> BpoiReturnCommonDTO<List<T>> importParseExcelLarge(String fullFileName, Class<T> excelParseClass,
                                                                  IExcelSaxRowRead rowRead, IExcelReadRowHandler handler,
                                                                  Integer maxErrHintCount) {
        if (fullFileName == null || "".equals(fullFileName)) {
            return new BpoiReturnCommonDTO(BpoiConstants.commonReturnStatus.FAIL.getValue(), "导入文件名为空");
        }
        try {
            InputStream fileInputStream = new FileInputStream(fullFileName);
            String excelType = fullFileName.substring(fullFileName.lastIndexOf(".") + 1);
            return importParseExcelLarge(excelType, fileInputStream, excelParseClass, rowRead, handler, maxErrHintCount);
        } catch (IOException e) {
            throw BpoiException.errWithDetail("文件读取失败", e.getMessage(), e);
        }
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
    public <T> BpoiReturnCommonDTO<List<T>> importParseExcelLarge(String excelType, InputStream fileInputStream,
                                                                  Class<T> excelParseClass, IExcelSaxRowRead rowRead,
                                                                  IExcelReadRowHandler handler, Integer maxErrHintCount) {
        if (!BpoiConstants.excelType.XLSX.getValue().equals(excelType)) {
            throw new BpoiAlertException("大文件解析只支持2007版本的Excel");
        }
        try {
            // 打开并读取Excel
            OPCPackage opcPackage = OPCPackage.open(fileInputStream);
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            SharedStringsTable sst = xssfReader.getSharedStringsTable();
            if (rowRead == null) {
                ExcelSaxImportParam importParam = new ExcelSaxImportParam(new StringJoiner("、"), maxErrHintCount, 0);
                rowRead = new ExcelSaxRowRead(excelParseClass, importParam, handler);
            }
            // 获取解析处理器
            XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
            ContentHandler saxHandler = new ExcelSheetHandler(sst, rowRead);
            parser.setContentHandler(saxHandler);
            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            if (sheets.hasNext()) {
                InputStream sheet = sheets.next();
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);
                sheet.close();
            }
            List<T> dataList = rowRead.getList();
            return new BpoiReturnCommonDTO<>(dataList);
        } catch (BpoiAlertException e) {
            return new BpoiReturnCommonDTO<>(BpoiConstants.commonReturnStatus.FAIL.getValue(), e.getMessage());
        } catch (Exception e) {
            throw BpoiException.errWithDetail("导入文件解析失败", e.getMessage(), e);
        }
    }

}
