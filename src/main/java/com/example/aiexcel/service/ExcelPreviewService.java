package com.example.aiexcel.service;

import com.example.aiexcel.service.excel.ExcelService;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

/**
 * Excel表格预览服务
 * 提供数据预览和可视化功能
 */
@Service
public class ExcelPreviewService {

    @Autowired
    private ExcelService excelService;

    private static final Logger logger = LoggerFactory.getLogger(ExcelPreviewService.class);

    /**
     * 获取Excel文件的预览数据
     * @param file Excel文件
     * @return 预览数据
     */
    public Map<String, Object> getExcelPreviewData(MultipartFile file) throws IOException {
        logger.info("Getting Excel preview data for file: {}", file.getOriginalFilename());

        // 加载工作簿
        Workbook workbook = excelService.loadWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0); // 默认获取第一个工作表

        // 获取数据
        Object[][] data = getSheetData(sheet);
        String[] headers = getHeaders(sheet);

        // 获取数据统计信息
        int totalRows = data.length;
        int totalColumns = headers.length;
        String sheetName = sheet.getSheetName();

        // 构建响应
        Map<String, Object> response = new HashMap<>();
        response.put("data", data);
        response.put("headers", headers);
        response.put("rowCount", totalRows);
        response.put("columnCount", totalColumns);
        response.put("sheetName", sheetName);
        response.put("success", true);

        logger.info("Successfully retrieved preview data for sheet: {}, rows: {}, cols: {}", 
                   sheetName, totalRows, totalColumns);

        return response;
    }

    /**
     * 获取工作表数据
     */
    private Object[][] getSheetData(Sheet sheet) {
        if (sheet == null) {
            return new Object[0][0];
        }

        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return new Object[0][0];
        }

        // 计算最大列数
        int maxCols = 0;
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getLastCellNum() > maxCols) {
                maxCols = row.getLastCellNum();
            }
        }

        // 创建数据数组
        Object[][] data = new Object[lastRowNum + 1][maxCols];

        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < maxCols; j++) {
                    Cell cell = row.getCell(j);
                    data[i][j] = getCellValueAsString(cell);
                }
            } else {
                // 空行填充空字符串
                for (int j = 0; j < maxCols; j++) {
                    data[i][j] = "";
                }
            }
        }

        return data;
    }

    /**
     * 获取表头
     */
    private String[] getHeaders(Sheet sheet) {
        if (sheet == null || sheet.getFirstRowNum() < 0) {
            return new String[0];
        }

        Row firstRow = sheet.getRow(sheet.getFirstRowNum());
        if (firstRow == null) {
            return new String[0];
        }

        int lastCellNum = firstRow.getLastCellNum();
        String[] headers = new String[lastCellNum];

        for (int i = 0; i < lastCellNum; i++) {
            Cell cell = firstRow.getCell(i);
            headers[i] = getCellValueAsString(cell);
            if (headers[i] == null || headers[i].trim().isEmpty()) {
                // 如果表头单元格为空，使用列字母作为表头
                headers[i] = getColumnLetter(i + 1);
            }
        }

        return headers;
    }

    /**
     * 获取单元格值的字符串表示
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        // 检查是否为整数
                        if (cell.getNumericCellValue() == Math.floor(cell.getNumericCellValue())) {
                            return String.valueOf((long) cell.getNumericCellValue());
                        } else {
                            return String.valueOf(cell.getNumericCellValue());
                        }
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    // 对于公式单元格，返回计算结果而不是公式本身
                    // 需要获取所在工作簿以创建公式计算器
                    Workbook workbook = cell.getSheet().getWorkbook();
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case STRING:
                            return cellValue.getStringValue();
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return DateUtil.getJavaDate(cellValue.getNumberValue()).toString();
                            } else {
                                double value = cellValue.getNumberValue();
                                // 检查是否实际为整数以避免 .0 小数
                                if (value == Math.floor(value)) {
                                    return String.valueOf((int) value);
                                } else {
                                    return String.valueOf(value);
                                }
                            }
                        case BOOLEAN:
                            return String.valueOf(cellValue.getBooleanValue());
                        case ERROR:
                            return "#ERROR!";
                        default:
                            return cellValue.formatAsString();
                    }
                case BLANK:
                    return "";
                case ERROR:
                    return "#ERROR!";
                default:
                    return cell.toString();
            }
        } catch (Exception e) {
            logger.warn("Error getting cell value for cell {}: {}", cell.getAddress(), e.getMessage());
            return cell.toString();
        }
    }

    /**
     * 将列索引转换为字母 (0->A, 1->B, ..., 26->AA, etc.)
     */
    private String getColumnLetter(int columnIndex) {
        if (columnIndex <= 0) {
            return "";
        }

        StringBuilder columnName = new StringBuilder();
        while (columnIndex > 0) {
            columnIndex--; // 调整为0基索引
            columnName.insert(0, (char) ('A' + columnIndex % 26));
            columnIndex = columnIndex / 26;
        }
        return columnName.toString();
    }

    /**
     * 获取单元格格式信息
     * @param file Excel文件
     * @param row 行索引
     * @param col 列索引
     * @return 格式信息
     */
    public Map<String, Object> getCellFormat(MultipartFile file, int row, int col) throws IOException {
        Workbook workbook = excelService.loadWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Row sheetRow = sheet.getRow(row);
        
        Map<String, Object> formatInfo = new HashMap<>();
        
        if (sheetRow != null) {
            Cell cell = sheetRow.getCell(col);
            if (cell != null) {
                CellStyle style = cell.getCellStyle();
                Font font = workbook.getFontAt(style.getFontIndex());
                
                formatInfo.put("backgroundColor", getBackgroundColor(style));
                formatInfo.put("foregroundColor", getForegroundColor(style));
                formatInfo.put("fontBold", font.getBold());
                formatInfo.put("fontItalic", font.getItalic());
                formatInfo.put("fontSize", font.getFontHeightInPoints());
                formatInfo.put("fontColor", getFontColor(font));
                formatInfo.put("horizontalAlignment", style.getAlignment());
                formatInfo.put("verticalAlignment", style.getVerticalAlignment());
                formatInfo.put("borderLeft", style.getBorderLeft());
                formatInfo.put("borderRight", style.getBorderRight());
                formatInfo.put("borderTop", style.getBorderTop());
                formatInfo.put("borderBottom", style.getBorderBottom());
            }
        }
        
        return formatInfo;
    }

    /**
     * 获取背景颜色
     */
    private String getBackgroundColor(CellStyle style) {
        // 简化处理，返回颜色索引或空字符串
        return style.getFillForegroundColor() > 0 ? String.valueOf(style.getFillForegroundColor()) : "";
    }

    /**
     * 获取前景颜色
     */
    private String getForegroundColor(CellStyle style) {
        // 简化处理，返回颜色索引或空字符串
        return style.getFillBackgroundColor() > 0 ? String.valueOf(style.getFillBackgroundColor()) : "";
    }

    /**
     * 获取字体颜色
     */
    private String getFontColor(Font font) {
        // 简化处理，返回颜色索引或空字符串
        short color = font.getColor();
        return color > 0 ? String.valueOf(color) : "";
    }

    /**
     * 获取单元格范围的格式信息（批量获取）
     * @param file Excel文件
     * @param startRow 起始行
     * @param startCol 起始列
     * @param endRow 结束行
     * @param endCol 结束列
     * @return 格式信息
     */
    public Map<String, Object> getBulkCellFormat(MultipartFile file, int startRow, int startCol, int endRow, int endCol) throws IOException {
        Workbook workbook = excelService.loadWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        Map<String, Object> result = new HashMap<>();
        Map<String, Object> formatData = new HashMap<>();

        // 遍历指定范围内的所有单元格
        for (int rowIdx = startRow; rowIdx <= endRow && rowIdx < sheet.getLastRowNum() + 1; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) continue;

            for (int colIdx = startCol; colIdx <= endCol && colIdx < row.getLastCellNum(); colIdx++) {
                Cell cell = row.getCell(colIdx);

                String cellKey = getCellKey(rowIdx, colIdx);
                Map<String, Object> cellFormat = new HashMap<>();

                if (cell != null) {
                    CellStyle style = cell.getCellStyle();
                    Font font = workbook.getFontAt(style.getFontIndex());

                    cellFormat.put("backgroundColor", getBackgroundColor(style));
                    cellFormat.put("foregroundColor", getForegroundColor(style));
                    cellFormat.put("fontBold", font.getBold());
                    cellFormat.put("fontItalic", font.getItalic());
                    cellFormat.put("fontSize", font.getFontHeightInPoints());
                    cellFormat.put("fontColor", getFontColor(font));
                    cellFormat.put("horizontalAlignment", style.getAlignment().name());
                    cellFormat.put("verticalAlignment", style.getVerticalAlignment().name());
                    cellFormat.put("borderLeft", style.getBorderLeft().name());
                    cellFormat.put("borderRight", style.getBorderRight().name());
                    cellFormat.put("borderTop", style.getBorderTop().name());
                    cellFormat.put("borderBottom", style.getBorderBottom().name());

                    // 获取数据类型
                    String dataType = getCellDataType(cell);
                    cellFormat.put("dataType", dataType);

                    // 获取原始值
                    cellFormat.put("value", getCellValueAsString(cell));
                } else {
                    // 空单元格的默认格式
                    cellFormat.put("backgroundColor", "");
                    cellFormat.put("foregroundColor", "");
                    cellFormat.put("fontBold", false);
                    cellFormat.put("fontItalic", false);
                    cellFormat.put("fontSize", 10);
                    cellFormat.put("fontColor", "");
                    cellFormat.put("horizontalAlignment", "GENERAL");
                    cellFormat.put("verticalAlignment", "BOTTOM");
                    cellFormat.put("borderLeft", "NONE");
                    cellFormat.put("borderRight", "NONE");
                    cellFormat.put("borderTop", "NONE");
                    cellFormat.put("borderBottom", "NONE");
                    cellFormat.put("dataType", "BLANK");
                    cellFormat.put("value", "");
                }

                formatData.put(cellKey, cellFormat);
            }
        }

        result.put("formatData", formatData);
        result.put("range", Map.of(
            "startRow", startRow,
            "startCol", startCol,
            "endRow", endRow,
            "endCol", endCol
        ));

        return result;
    }

    /**
     * 获取单元格键值（用于标识特定单元格）
     */
    private String getCellKey(int row, int col) {
        return row + "_" + col;
    }

    /**
     * 获取单元格数据类型
     */
    private String getCellDataType(Cell cell) {
        if (cell == null) {
            return "BLANK";
        }

        switch (cell.getCellType()) {
            case STRING:
                return "STRING";
            case NUMERIC:
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    return "DATE";
                } else {
                    return "NUMBER";
                }
            case BOOLEAN:
                return "BOOLEAN";
            case FORMULA:
                return "FORMULA";
            case BLANK:
                return "BLANK";
            case ERROR:
                return "ERROR";
            default:
                return "UNKNOWN";
        }
    }
}