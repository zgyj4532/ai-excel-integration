package com.example.aiexcel.service;

import com.example.aiexcel.service.excel.ExcelService;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * AI Excel命令解析器
 * 用于解析AI返回的Excel操作指令并执行相应操作
 */
@Service
public class AiExcelCommandParser {

    private final ExcelService excelService;
    private static final Logger logger = LoggerFactory.getLogger(AiExcelCommandParser.class);

    @Autowired
    public AiExcelCommandParser(ExcelService excelService) {
        this.excelService = excelService;
    }

    /**
     * 解析AI响应中的Excel操作指令并执行
     *
     * @param workbook  Excel工作簿
     * @param aiResponse AI响应文本
     * @return 执行结果
     */
    public List<CommandResult> parseAndExecuteCommands(Workbook workbook, String aiResponse) {
        List<CommandResult> results = new ArrayList<>();

        if (aiResponse == null || aiResponse.trim().isEmpty()) {
            logger.warn("AI response is null or empty, no commands to execute");
            return results;
        }

        logger.info("Starting to parse AI response for Excel commands: {}", aiResponse);

        // 支持的命令格式
        // [SET_CELL:A1:New Value] - 设置单元格A1为'New Value'
        // [INSERT_ROW:3:value1,value2,value3] - 在第3行插入值
        // [INSERT_COLUMN:2:value1,value2,value3] - 在第2列插入值
        // [DELETE_ROW:5] - 删除第5行
        // [DELETE_COLUMN:1] - 删除第1列
        // [APPLY_FORMULA:A1:B1+C1] - 在A1应用公式B1+C1

        // 解析设置单元格值的命令
        Pattern setCellPattern = Pattern.compile("\\[SET_CELL:([A-Z]+\\d+):(.+?)\\]");
        Matcher setCellMatcher = setCellPattern.matcher(aiResponse);

        while (setCellMatcher.find()) {
            String cellRef = setCellMatcher.group(1);
            String value = setCellMatcher.group(2);

            try {
                // 验证单元格引用格式
                if (!isValidCellReference(cellRef)) {
                    logger.error("Invalid cell reference format: {}", cellRef);
                    results.add(new CommandResult(false, "SET_CELL", cellRef + "=" + value, "Invalid cell reference format: " + cellRef));
                    continue;
                }

                // 解析单元格引用 (如 A1 -> col=0, row=0)
                CellReference ref = parseCellReference(cellRef);
                excelService.updateCell(workbook, workbook.getSheetName(0), ref.row, ref.col, value);
                logger.info("Successfully set cell {} to {}", cellRef, value);
                results.add(new CommandResult(true, "SET_CELL", cellRef + "=" + value, "Successfully set cell " + cellRef + " to " + value));
            } catch (Exception e) {
                logger.error("Error setting cell {}: {}", cellRef, e.getMessage(), e);
                results.add(new CommandResult(false, "SET_CELL", cellRef + "=" + value, "Error setting cell " + cellRef + ": " + e.getMessage()));
            }
        }

        // 解析插入行命令
        Pattern insertRowPattern = Pattern.compile("\\[INSERT_ROW:(\\d+):(.+?)\\]");
        Matcher insertRowMatcher = insertRowPattern.matcher(aiResponse);

        while (insertRowMatcher.find()) {
            String rowIndexStr = insertRowMatcher.group(1);
            String valuesStr = insertRowMatcher.group(2);

            try {
                int rowIndex = Integer.parseInt(rowIndexStr);
                if (rowIndex < 0) {
                    logger.error("Invalid row index: {}", rowIndex);
                    results.add(new CommandResult(false, "INSERT_ROW", rowIndexStr + ":" + valuesStr, "Invalid row index: " + rowIndex));
                    continue;
                }

                String[] values = valuesStr.split(",", -1); // 使用-1以保留尾随空值

                // 在指定行插入数据
                insertRow(workbook, rowIndex, values);
                logger.info("Successfully inserted row at {}", rowIndex);
                results.add(new CommandResult(true, "INSERT_ROW", rowIndexStr + ":" + valuesStr, "Successfully inserted row at " + rowIndex));
            } catch (NumberFormatException e) {
                logger.error("Invalid row index format: {}", rowIndexStr);
                results.add(new CommandResult(false, "INSERT_ROW", rowIndexStr + ":" + valuesStr, "Invalid row index format: " + rowIndexStr));
            } catch (Exception e) {
                logger.error("Error inserting row at {}: {}", rowIndexStr, e.getMessage(), e);
                results.add(new CommandResult(false, "INSERT_ROW", rowIndexStr + ":" + valuesStr, "Error inserting row at " + rowIndexStr + ": " + e.getMessage()));
            }
        }

        // 解析插入列命令
        Pattern insertColPattern = Pattern.compile("\\[INSERT_COLUMN:(\\d+):(.+?)\\]");
        Matcher insertColMatcher = insertColPattern.matcher(aiResponse);

        while (insertColMatcher.find()) {
            String colIndexStr = insertColMatcher.group(1);
            String valuesStr = insertColMatcher.group(2);

            try {
                int colIndex = Integer.parseInt(colIndexStr);
                if (colIndex < 0) {
                    logger.error("Invalid column index: {}", colIndex);
                    results.add(new CommandResult(false, "INSERT_COLUMN", colIndexStr + ":" + valuesStr, "Invalid column index: " + colIndex));
                    continue;
                }

                String[] values = valuesStr.split(",", -1);

                // 在指定列插入数据
                insertColumn(workbook, colIndex, values);
                logger.info("Successfully inserted column at {}", colIndex);
                results.add(new CommandResult(true, "INSERT_COLUMN", colIndexStr + ":" + valuesStr, "Successfully inserted column at " + colIndex));
            } catch (NumberFormatException e) {
                logger.error("Invalid column index format: {}", colIndexStr);
                results.add(new CommandResult(false, "INSERT_COLUMN", colIndexStr + ":" + valuesStr, "Invalid column index format: " + colIndexStr));
            } catch (Exception e) {
                logger.error("Error inserting column at {}: {}", colIndexStr, e.getMessage(), e);
                results.add(new CommandResult(false, "INSERT_COLUMN", colIndexStr + ":" + valuesStr, "Error inserting column at " + colIndexStr + ": " + e.getMessage()));
            }
        }

        // 解析删除行命令
        Pattern deleteRowPattern = Pattern.compile("\\[DELETE_ROW:(\\d+)\\]");
        Matcher deleteRowMatcher = deleteRowPattern.matcher(aiResponse);

        while (deleteRowMatcher.find()) {
            String rowIndexStr = deleteRowMatcher.group(1);

            try {
                int rowIndex = Integer.parseInt(rowIndexStr);
                if (rowIndex < 0) {
                    logger.error("Invalid row index: {}", rowIndex);
                    results.add(new CommandResult(false, "DELETE_ROW", rowIndexStr, "Invalid row index: " + rowIndex));
                    continue;
                }

                // 删除指定行
                deleteRow(workbook, rowIndex);
                logger.info("Successfully deleted row {}", rowIndex);
                results.add(new CommandResult(true, "DELETE_ROW", rowIndexStr, "Successfully deleted row " + rowIndex));
            } catch (NumberFormatException e) {
                logger.error("Invalid row index format: {}", rowIndexStr);
                results.add(new CommandResult(false, "DELETE_ROW", rowIndexStr, "Invalid row index format: " + rowIndexStr));
            } catch (Exception e) {
                logger.error("Error deleting row {}: {}", rowIndexStr, e.getMessage(), e);
                results.add(new CommandResult(false, "DELETE_ROW", rowIndexStr, "Error deleting row " + rowIndexStr + ": " + e.getMessage()));
            }
        }

        // 解析删除列命令
        Pattern deleteColPattern = Pattern.compile("\\[DELETE_COLUMN:(\\d+)\\]");
        Matcher deleteColMatcher = deleteColPattern.matcher(aiResponse);

        while (deleteColMatcher.find()) {
            String colIndexStr = deleteColMatcher.group(1);

            try {
                int colIndex = Integer.parseInt(colIndexStr);
                if (colIndex < 0) {
                    logger.error("Invalid column index: {}", colIndex);
                    results.add(new CommandResult(false, "DELETE_COLUMN", colIndexStr, "Invalid column index: " + colIndex));
                    continue;
                }

                // 删除指定列
                deleteColumn(workbook, colIndex);
                logger.info("Successfully deleted column {}", colIndex);
                results.add(new CommandResult(true, "DELETE_COLUMN", colIndexStr, "Successfully deleted column " + colIndex));
            } catch (NumberFormatException e) {
                logger.error("Invalid column index format: {}", colIndexStr);
                results.add(new CommandResult(false, "DELETE_COLUMN", colIndexStr, "Invalid column index format: " + colIndexStr));
            } catch (Exception e) {
                logger.error("Error deleting column {}: {}", colIndexStr, e.getMessage(), e);
                results.add(new CommandResult(false, "DELETE_COLUMN", colIndexStr, "Error deleting column " + colIndexStr + ": " + e.getMessage()));
            }
        }

        // 解析应用公式命令 - 缓存公式待后续计算
        Pattern applyFormulaPattern = Pattern.compile("\\[APPLY_FORMULA:([A-Z]+\\d+):(.+?)\\]");
        Matcher applyFormulaMatcher = applyFormulaPattern.matcher(aiResponse);

        // 创建缓存队列存储待计算的公式
        java.util.List<FormulaTask> formulaTasks = new java.util.ArrayList<>();

        while (applyFormulaMatcher.find()) {
            String cellRef = applyFormulaMatcher.group(1);
            String formula = applyFormulaMatcher.group(2);

            try {
                // 验证单元格引用格式
                if (!isValidCellReference(cellRef)) {
                    logger.error("Invalid cell reference format: {}", cellRef);
                    results.add(new CommandResult(false, "APPLY_FORMULA", cellRef + "=" + formula, "Invalid cell reference format: " + cellRef));
                    continue;
                }

                CellReference ref = parseCellReference(cellRef);
                // 缓存公式任务，待其他操作完成后统一计算
                FormulaTask task = new FormulaTask(cellRef, ref.row, ref.col, formula, workbook.getSheetName(0));
                formulaTasks.add(task);
                logger.info("Queued formula calculation for cell {}: {}", cellRef, formula);
                results.add(new CommandResult(true, "APPLY_FORMULA", cellRef + "=" + formula, "Formula calculation queued for cell " + cellRef));
            } catch (Exception e) {
                logger.error("Error queuing formula for cell {}: {}", cellRef, e.getMessage(), e);
                results.add(new CommandResult(false, "APPLY_FORMULA", cellRef + "=" + formula, "Error queuing formula for cell " + cellRef + ": " + e.getMessage()));
            }
        }

        // 在处理完其他命令后，统一计算所有公式
        for (FormulaTask task : formulaTasks) {
            try {
                Object calculatedResult = calculateFormulaResult(workbook, task.sheetName, task.row, task.col, task.formula);
                excelService.updateCell(workbook, task.sheetName, task.row, task.col, calculatedResult);
                logger.info("Successfully calculated and set result {} to cell {}", calculatedResult, task.cellRef);

                // 更新结果列表中的消息
                for (int i = 0; i < results.size(); i++) {
                    CommandResult result = results.get(i);
                    if (result.getCommandType().equals("APPLY_FORMULA") &&
                        result.getCommandParams().startsWith(task.cellRef + "=")) {
                        results.set(i, new CommandResult(true, "APPLY_FORMULA",
                            task.cellRef + "=" + task.formula,
                            "Successfully calculated and set result " + calculatedResult + " to cell " + task.cellRef));
                        break;
                    }
                }
            } catch (Exception e) {
                logger.error("Error processing queued formula for cell {}: {}", task.cellRef, e.getMessage(), e);

                // 更新结果列表中的消息
                for (int i = 0; i < results.size(); i++) {
                    CommandResult result = results.get(i);
                    if (result.getCommandType().equals("APPLY_FORMULA") &&
                        result.getCommandParams().startsWith(task.cellRef + "=")) {
                        results.set(i, new CommandResult(false, "APPLY_FORMULA",
                            task.cellRef + "=" + task.formula,
                            "Error processing formula for cell " + task.cellRef + ": " + e.getMessage()));
                        break;
                    }
                }
            }
        }

        logger.info("Completed parsing AI response, processed {} commands", results.size());
        return results;
    }

    /**
     * 验证单元格引用格式是否正确
     *
     * @param cellRef 单元格引用 (如 A1, B2, Z10)
     * @return 是否有效
     */
    private boolean isValidCellReference(String cellRef) {
        if (cellRef == null || cellRef.isEmpty()) {
            return false;
        }

        // 正则表达式验证单元格引用格式 (如 A1, B2, Z10, AA1, AB2, etc.)
        Pattern cellPattern = Pattern.compile("[A-Z]+\\d+");
        return cellPattern.matcher(cellRef.toUpperCase()).matches();
    }
    
    /**
     * 解析单元格引用 (如 A1 -> col=0, row=0)
     */
    private CellReference parseCellReference(String cellRef) throws IllegalArgumentException {
        if (cellRef == null || cellRef.isEmpty()) {
            throw new IllegalArgumentException("Cell reference cannot be null or empty");
        }

        StringBuilder colPart = new StringBuilder();
        StringBuilder rowPart = new StringBuilder();

        for (int i = 0; i < cellRef.length(); i++) {
            char c = cellRef.charAt(i);
            if (Character.isLetter(c)) {
                colPart.append(Character.toUpperCase(c));  // 统一转为大写
            } else if (Character.isDigit(c)) {
                rowPart.append(c);
            } else {
                throw new IllegalArgumentException("Invalid character in cell reference: " + c);
            }
        }

        if (colPart.length() == 0 || rowPart.length() == 0) {
            throw new IllegalArgumentException("Invalid cell reference format: " + cellRef);
        }

        int col = columnToNumber(colPart.toString());
        int row;
        try {
            row = Integer.parseInt(rowPart.toString()) - 1; // Excel行号从1开始，数组索引从0开始
            if (row < 0) {
                throw new IllegalArgumentException("Row index out of range: " + rowPart.toString());
            }
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("Invalid row number in cell reference: " + rowPart.toString());
        }

        return new CellReference(row, col);
    }

    /**
     * 将列字母转换为数字 (A=0, B=1, ..., Z=25, AA=26, ...)
     */
    private int columnToNumber(String columnName) {
        if (columnName == null || columnName.isEmpty()) {
            throw new IllegalArgumentException("Column name cannot be null or empty");
        }

        int result = 0;
        for (int i = 0; i < columnName.length(); i++) {
            char c = columnName.charAt(i);
            if (c < 'A' || c > 'Z') {
                throw new IllegalArgumentException("Invalid column letter: " + c);
            }
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1; // 转换为0基索引
    }

    /**
     * 插入一行数据
     */
    private void insertRow(Workbook workbook, int rowIndex, String[] values) {
        if (workbook == null) {
            logger.error("Workbook cannot be null");
            throw new IllegalArgumentException("Workbook cannot be null");
        }

        if (values == null) {
            logger.warn("Values array is null, inserting empty row at index {}", rowIndex);
            values = new String[0];
        }

        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            logger.error("Sheet at index 0 is null");
            throw new IllegalStateException("Sheet at index 0 is null");
        }

        try {
            // 移动现有行，为新行腾出空间
            sheet.shiftRows(rowIndex, sheet.getLastRowNum(), 1);

            // 创建新行并设置值
            org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowIndex);
            for (int i = 0; i < values.length; i++) {
                org.apache.poi.ss.usermodel.Cell cell = row.createCell(i);
                if (values[i] != null) {
                    try {
                        // 尝试解析为数字
                        double numericValue = Double.parseDouble(values[i].trim());
                        cell.setCellValue(numericValue);
                    } catch (NumberFormatException e) {
                        // 非数字则作为字符串
                        cell.setCellValue(values[i].trim());
                    }
                } else {
                    cell.setCellValue("");
                }
            }
        } catch (Exception e) {
            logger.error("Error inserting row at {}: {}", rowIndex, e.getMessage(), e);
            throw new RuntimeException("Error inserting row at " + rowIndex + ": " + e.getMessage(), e);
        }
    }

    /**
     * 插入一列数据
     */
    private void insertColumn(Workbook workbook, int colIndex, String[] values) {
        if (workbook == null) {
            logger.error("Workbook cannot be null");
            throw new IllegalArgumentException("Workbook cannot be null");
        }

        if (values == null) {
            logger.warn("Values array is null, inserting empty column at index {}", colIndex);
            values = new String[0];
        }

        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            logger.error("Sheet at index 0 is null");
            throw new IllegalStateException("Sheet at index 0 is null");
        }

        try {
            // 为每一行插入新列
            for (int i = 0; i < values.length && i < sheet.getLastRowNum() + 1; i++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(i);
                if (row == null) {
                    row = sheet.createRow(i);
                }

                // 移动现有单元格
                for (int j = Math.max(row.getLastCellNum(), colIndex); j > colIndex; j--) {
                    org.apache.poi.ss.usermodel.Cell srcCell = row.getCell(j - 1);
                    if (srcCell != null) {
                        org.apache.poi.ss.usermodel.Cell targetCell = row.getCell(j);
                        if (targetCell == null) {
                            targetCell = row.createCell(j);
                        }
                        // 复制单元格内容
                        copyCellContent(srcCell, targetCell);
                    }
                }

                // 设置新的单元格值
                org.apache.poi.ss.usermodel.Cell newCell = row.getCell(colIndex);
                if (newCell == null) {
                    newCell = row.createCell(colIndex);
                }
                if (values[i] != null) {
                    try {
                        // 尝试解析为数字
                        double numericValue = Double.parseDouble(values[i].trim());
                        newCell.setCellValue(numericValue);
                    } catch (NumberFormatException e) {
                        // 非数字则作为字符串
                        newCell.setCellValue(values[i].trim());
                    }
                } else {
                    newCell.setCellValue("");
                }
            }
        } catch (Exception e) {
            logger.error("Error inserting column at {}: {}", colIndex, e.getMessage(), e);
            throw new RuntimeException("Error inserting column at " + colIndex + ": " + e.getMessage(), e);
        }
    }

    /**
     * 删除指定行
     */
    private void deleteRow(Workbook workbook, int rowIndex) {
        if (workbook == null) {
            logger.error("Workbook cannot be null");
            throw new IllegalArgumentException("Workbook cannot be null");
        }

        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            logger.error("Sheet at index 0 is null");
            throw new IllegalStateException("Sheet at index 0 is null");
        }

        try {
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIndex);
            if (row != null) {
                sheet.removeRow(row);
            }
            // 重新整理行索引
            if (rowIndex < sheet.getLastRowNum()) {
                sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), -1);
            }
        } catch (Exception e) {
            logger.error("Error deleting row at {}: {}", rowIndex, e.getMessage(), e);
            throw new RuntimeException("Error deleting row at " + rowIndex + ": " + e.getMessage(), e);
        }
    }

    /**
     * 删除指定列
     */
    private void deleteColumn(Workbook workbook, int colIndex) {
        if (workbook == null) {
            logger.error("Workbook cannot be null");
            throw new IllegalArgumentException("Workbook cannot be null");
        }

        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            logger.error("Sheet at index 0 is null");
            throw new IllegalStateException("Sheet at index 0 is null");
        }

        try {
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(i);
                if (row != null) {
                    org.apache.poi.ss.usermodel.Cell cell = row.getCell(colIndex);
                    if (cell != null) {
                        row.removeCell(cell);
                    }

                    // 移动后续列的单元格
                    for (int j = colIndex; j < row.getLastCellNum(); j++) {
                        org.apache.poi.ss.usermodel.Cell srcCell = row.getCell(j + 1);
                        if (srcCell != null) {
                            org.apache.poi.ss.usermodel.Cell targetCell = row.getCell(j);
                            if (targetCell == null) {
                                targetCell = row.createCell(j);
                            }
                            copyCellContent(srcCell, targetCell);
                            row.removeCell(srcCell);
                        }
                    }
                }
            }
        } catch (Exception e) {
            logger.error("Error deleting column at {}: {}", colIndex, e.getMessage(), e);
            throw new RuntimeException("Error deleting column at " + colIndex + ": " + e.getMessage(), e);
        }
    }

    /**
     * 复制单元格内容
     */
    private void copyCellContent(org.apache.poi.ss.usermodel.Cell src, org.apache.poi.ss.usermodel.Cell target) {
        if (src == null || target == null) {
            logger.warn("Source or target cell is null in copyCellContent");
            return;
        }

        try {
            switch (src.getCellType()) {
                case STRING:
                    target.setCellValue(src.getStringCellValue());
                    break;
                case NUMERIC:
                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(src)) {
                        target.setCellValue(src.getDateCellValue());
                    } else {
                        target.setCellValue(src.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    target.setCellValue(src.getBooleanCellValue());
                    break;
                case FORMULA:
                    target.setCellFormula(src.getCellFormula());
                    break;
                case BLANK:
                    target.setBlank();
                    break;
                case ERROR:
                    target.setCellErrorValue(src.getErrorCellValue());
                    break;
                default:
                    target.setCellValue(src.toString());
            }
        } catch (Exception e) {
            logger.error("Error copying cell content from {} to {}: {}", src.getAddress(), target.getAddress(), e.getMessage(), e);
            target.setCellValue(src.toString()); // Fallback to string conversion
        }
    }
    
    /**
     * 内部类表示单元格引用
     */
    private static class CellReference {
        int row;
        int col;
        
        CellReference(int row, int col) {
            this.row = row;
            this.col = col;
        }
    }
    
    /**
     * 公式任务类 - 用于缓存待计算的公式
     */
    private static class FormulaTask {
        String cellRef;
        int row;
        int col;
        String formula;
        String sheetName;

        FormulaTask(String cellRef, int row, int col, String formula, String sheetName) {
            this.cellRef = cellRef;
            this.row = row;
            this.col = col;
            this.formula = formula;
            this.sheetName = sheetName;
        }
    }

    /**
     * 命令执行结果类
     */
    public static class CommandResult {
        private boolean success;
        private String commandType;
        private String commandParams;
        private String message;

        public CommandResult(boolean success, String commandType, String commandParams, String message) {
            this.success = success;
            this.commandType = commandType;
            this.commandParams = commandParams;
            this.message = message;
        }

        // Getters and setters
        public boolean isSuccess() { return success; }
        public void setSuccess(boolean success) { this.success = success; }

        public String getCommandType() { return commandType; }
        public void setCommandType(String commandType) { this.commandType = commandType; }

        public String getCommandParams() { return commandParams; }
        public void setCommandParams(String commandParams) { this.commandParams = commandParams; }

        public String getMessage() { return message; }
        public void setMessage(String message) { this.message = message; }
    }

    /**
     * 计算公式结果
     * @param workbook Excel工作簿
     * @param sheetName 工作表名称
     * @param row 行索引
     * @param col 列索引
     * @param formula 公式
     * @return 计算结果
     */
    private Object calculateFormulaResult(Workbook workbook, String sheetName, int row, int col, String formula) {
        if (formula == null || formula.trim().isEmpty()) {
            return null;
        }

        // 去除公式的等号前缀
        String cleanFormula = formula.trim();
        if (cleanFormula.startsWith("=")) {
            cleanFormula = cleanFormula.substring(1);
        }

        // 根据AI的自然语言命令，我们可能需要处理不同类型的公式
        // 最常见的是引用其他单元格的运算，如 A1+B1, C1*D1 等
        try {
            // 对于简单的加法操作（如 A1+B1, B2+D2），我们手动解析并计算
            if (cleanFormula.contains("+")) {
                return calculateAddition(workbook, sheetName, cleanFormula);
            } else if (cleanFormula.contains("-") && !cleanFormula.startsWith("-")) {
                return calculateSubtraction(workbook, sheetName, cleanFormula);
            } else if (cleanFormula.contains("*")) {
                return calculateMultiplication(workbook, sheetName, cleanFormula);
            } else if (cleanFormula.contains("/")) {
                return calculateDivision(workbook, sheetName, cleanFormula);
            } else {
                // 如果是单个单元格引用或常量，直接返回值
                return getCellValue(workbook, sheetName, cleanFormula);
            }
        } catch (Exception e) {
            logger.error("Error calculating formula '{}': {}", formula, e.getMessage(), e);
            return formula; // 返回原始公式如果无法计算
        }
    }

    /**
     * 计算加法
     */
    private Object calculateAddition(Workbook workbook, String sheetName, String formula) {
        String[] parts = formula.split("\\+");
        if (parts.length != 2) {
            throw new IllegalArgumentException("Invalid addition formula: " + formula);
        }

        Object val1 = getCellValue(workbook, sheetName, parts[0].trim());
        Object val2 = getCellValue(workbook, sheetName, parts[1].trim());

        // 尝试转换为数字并相加
        try {
            double num1 = convertToNumber(val1);
            double num2 = convertToNumber(val2);
            return num1 + num2;
        } catch (NumberFormatException e) {
            // 如果不能转换为数字，按字符串处理
            return String.valueOf(val1) + String.valueOf(val2);
        }
    }

    /**
     * 计算减法
     */
    private Object calculateSubtraction(Workbook workbook, String sheetName, String formula) {
        String[] parts = formula.split("-", 2); // 限制分割次数以处理负数
        if (parts.length != 2) {
            throw new IllegalArgumentException("Invalid subtraction formula: " + formula);
        }

        Object val1 = getCellValue(workbook, sheetName, parts[0].trim());
        Object val2 = getCellValue(workbook, sheetName, parts[1].trim());

        // 尝试转换为数字并相减
        try {
            double num1 = convertToNumber(val1);
            double num2 = convertToNumber(val2);
            return num1 - num2;
        } catch (NumberFormatException e) {
            // 如果不能转换为数字，返回错误信息
            throw new NumberFormatException("Cannot perform subtraction on non-numeric values: " + val1 + " - " + val2);
        }
    }

    /**
     * 计算乘法
     */
    private Object calculateMultiplication(Workbook workbook, String sheetName, String formula) {
        String[] parts = formula.split("\\*");
        if (parts.length != 2) {
            throw new IllegalArgumentException("Invalid multiplication formula: " + formula);
        }

        Object val1 = getCellValue(workbook, sheetName, parts[0].trim());
        Object val2 = getCellValue(workbook, sheetName, parts[1].trim());

        // 尝试转换为数字并相乘
        try {
            double num1 = convertToNumber(val1);
            double num2 = convertToNumber(val2);
            return num1 * num2;
        } catch (NumberFormatException e) {
            // 如果不能转换为数字，返回错误信息
            throw new NumberFormatException("Cannot perform multiplication on non-numeric values: " + val1 + " * " + val2);
        }
    }

    /**
     * 计算除法
     */
    private Object calculateDivision(Workbook workbook, String sheetName, String formula) {
        String[] parts = formula.split("/");
        if (parts.length != 2) {
            throw new IllegalArgumentException("Invalid division formula: " + formula);
        }

        Object val1 = getCellValue(workbook, sheetName, parts[0].trim());
        Object val2 = getCellValue(workbook, sheetName, parts[1].trim());

        // 尝试转换为数字并相除
        try {
            double num1 = convertToNumber(val1);
            double num2 = convertToNumber(val2);
            if (num2 == 0) {
                throw new ArithmeticException("Division by zero");
            }
            return num1 / num2;
        } catch (NumberFormatException e) {
            // 如果不能转换为数字，返回错误信息
            throw new NumberFormatException("Cannot perform division on non-numeric values: " + val1 + " / " + val2);
        }
    }

    /**
     * 获取单元格值
     */
    private Object getCellValue(Workbook workbook, String sheetName, String cellRef) {
        // 如果cellRef 是一个单元格引用 (如 A1, B2)
        if (isValidCellReference(cellRef)) {
            CellReference ref = parseCellReference(cellRef);
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet(sheetName);
            if (sheet != null) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(ref.row);
                if (row != null) {
                    org.apache.poi.ss.usermodel.Cell cell = row.getCell(ref.col);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                return cell.getStringCellValue();
                            case NUMERIC:
                                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                    return cell.getDateCellValue();
                                } else {
                                    return cell.getNumericCellValue();
                                }
                            case BOOLEAN:
                                return cell.getBooleanCellValue();
                            case FORMULA:
                                // 对公式单元格进行计算
                                org.apache.poi.ss.usermodel.FormulaEvaluator evaluator =
                                    workbook.getCreationHelper().createFormulaEvaluator();
                                org.apache.poi.ss.usermodel.CellValue cellValue = evaluator.evaluate(cell);
                                switch (cellValue.getCellType()) {
                                    case STRING:
                                        return cellValue.getStringValue();
                                    case NUMERIC:
                                        return cellValue.getNumberValue();
                                    case BOOLEAN:
                                        return cellValue.getBooleanValue();
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
                    }
                }
            }
        } else {
            // 如果不是单元格引用，可能是常量值
            // 尝试解析为数字，如果失败则作为字符串返回
            try {
                return Double.parseDouble(cellRef);
            } catch (NumberFormatException e) {
                return cellRef; // 返回原始字符串
            }
        }

        return null;
    }

    /**
     * 转换为数字
     */
    private double convertToNumber(Object value) {
        if (value == null) {
            return 0.0;
        }

        if (value instanceof Number) {
            return ((Number) value).doubleValue();
        }

        if (value instanceof String) {
            String str = ((String) value).trim();
            if (str.isEmpty()) {
                return 0.0;
            }
            return Double.parseDouble(str);
        }

        // 对其他类型尝试转换
        return Double.parseDouble(value.toString());
    }
}