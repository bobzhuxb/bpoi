# Java解析Excel工具BPOI

基于POI 3.17版本开发的Excel导入导出便利化工具

### 快速开始

#### 导出Excel（HTTP导出下载）
    public static void simpleHttpExport1(HttpServletResponse response) {
        ExcelService.exportExcel(
                response,
                new ExcelExportDTO()
                        // Excel文件名，只有一个sheet时，也作为sheet名（不包括后缀，默认只导出xlsx）
                        .setFileName("xx报表")
                        // 列标题（只支持单行的列）
                        .setTitleList(Arrays.asList(
                                new ExcelTitleDTO("name", "名称"),
                                new ExcelTitleDTO("content", "内容")
                        ))
                        // 列表数据（Bean方式）
                        .setDataList(Arrays.asList(
                                new DemoExportDTO("苹果", "描述1"),
                                new DemoExportDTO("香蕉", "描述2"),
                                new DemoExportDTO("草莓", "描述3"),
                                new DemoExportDTO("西瓜", "描述4")
                        ))
        );
    }

#### 导出Excel（指定本地目录导出）
    public static void simpleHttpExport2() {
        ExcelService.exportExcel(
                null,
                new ExcelExportDTO()
                        // 绝对路径
                        .setBaseLocation("D:\\Test")
                        // 在上述绝对路径下，带相对路径的文件名（不包括后缀，默认只导出xlsx）
                        .setFileName("\\test0\\xx报表")
                        // 列标题（只支持单行的列）
                        .setTitleList(Arrays.asList(
                                new ExcelTitleDTO("index", "序号"),
                                new ExcelTitleDTO("name", "名称"),
                                new ExcelTitleDTO("content", "内容")
                        ))
                        // 列表数据（Map方式）
                        .setDataList(Arrays.asList(
                                new HashMap() {{put("index", 1); put("name", "苹果"); put("content", "描述1");}},
                                new HashMap() {{put("index", 2); put("name", "香蕉"); put("content", "描述2");}},
                                new HashMap() {{put("index", 3); put("name", "草莓"); put("content", "描述3");}},
                                new HashMap() {{put("index", 4); put("name", "西瓜"); put("content", "描述4");}}
                        ))
        );
    }

#### 导出Excel（稍复杂的导出）
    public static void complexExport() {
        // 列表数据内容
        List<Map<String, Object>> dataList = Arrays.asList(
                new HashMap() {{put("index", 1); put("name", "苹果"); put("content", "描述1"); put("detail", "明细1"); put("result", "结果1");}},
                new HashMap() {{put("index", 2); put("name", "香蕉"); put("content", "描述2"); put("detail", "明细2"); put("result", "结果2");}},
                new HashMap() {{put("index", 3); put("name", "草莓"); put("content", "描述3"); put("detail", "明细3"); put("result", "结果3");}},
                new HashMap() {{put("index", 4); put("name", "西瓜"); put("content", "描述4"); put("detail", "明细4"); put("result", "结果4");}}
        );
        // 合并单元格的行号和列号需精确计算
        int afterDataRow = dataList.size() + 3;
        // Excel导出
        ExcelService.exportExcel(
                null,
                new ExcelExportDTO()
                        .setBaseLocation("D:\\Test")
                        .setFileName("\\test0\\yy报表")
                        .setMaxColumn(5)
                        .setTableStartRow(2)
                        .setBeforeDataCellList(Arrays.asList(
                                // ExcelCellMainTitleBuilder是内置的
                                BpoiExcelUtil.buildCell(new ExcelCellMainTitleBuilder(0, 0, "yy报表")),
                                BpoiExcelUtil.buildCell(new ExcelCellBorderKeyBuilder(1, 0, "企业名称")),
                                BpoiExcelUtil.buildCell(new ExcelCellBorderValueBuilder(1, 2, "xx企业"))
                        ))
                        .setAfterDataCellList(Arrays.asList(
                                BpoiExcelUtil.buildCell(new ExcelCellBorderKeyBuilder(0, 0, "其他情况")),
                                BpoiExcelUtil.buildCell(new ExcelCellBorderValueBuilder(0, 2, "xxxxxx")),
                                BpoiExcelUtil.buildCell(new ExcelCellBorderKeyBuilder(1, 0, "报表生成时间")),
                                BpoiExcelUtil.buildCell(new ExcelCellBorderValueBuilder(1, 2, "2020-01-01 12:00:00"))
                        ))
                        .setTitleList(Arrays.asList(
                                new ExcelTitleDTO("index", "序号"),
                                new ExcelTitleDTO("name", "名称"),
                                new ExcelTitleDTO("content", "内容"),
                                new ExcelTitleDTO("detail", "明细"),
                                new ExcelTitleDTO("result", "结果")))
                        .setDataList(dataList)
                        .setDataSpecialStyleMap(new HashMap<Integer, ExcelCellDTO>() {{
                            put(2, BpoiExcelUtil.buildCell(new ExcelCellValueLeftWrapBuilder()));
                            put(3, BpoiExcelUtil.buildCell(new ExcelCellValueLeftWrapBuilder()));
                        }})
                        .setCellRangeList(Arrays.asList(
                                new ExcelCellRangeDTO(0, 0, 0, 4),
                                new ExcelCellRangeDTO(1, 1, 0, 1).setAllBorder(BorderStyle.THIN),
                                new ExcelCellRangeDTO(1, 1, 2, 4).setAllBorder(BorderStyle.THIN),
                                new ExcelCellRangeDTO(afterDataRow, afterDataRow, 0, 1).setAllBorder(BorderStyle.THIN),
                                new ExcelCellRangeDTO(afterDataRow, afterDataRow, 2, 4).setAllBorder(BorderStyle.THIN),
                                new ExcelCellRangeDTO(afterDataRow + 1, afterDataRow + 1, 0, 1).setAllBorder(BorderStyle.THIN),
                                new ExcelCellRangeDTO(afterDataRow + 1, afterDataRow + 1, 2, 4).setAllBorder(BorderStyle.THIN))));
    }

#### 导入Excel（只支持单Sheet，单行标题）
    public static void simpleImport() throws Exception {
        InputStream fileInputStream = new FileInputStream("D:\\Test\\test0\\xx报表.xlsx");
        BpoiReturnCommonDTO<List<DemoImportDTO>> rtnData = ExcelService.importParseExcel(
                BpoiConstants.excelType.XLSX.getValue(), fileInputStream, DemoImportDTO.class, null);
        if (BpoiConstants.commonReturnStatus.SUCCESS.getValue().equals(rtnData.getResultCode())) {
            List<DemoImportDTO> dataList = rtnData.getData();
            dataList.forEach(data -> System.out.println(data.toString()));
        }
    }
