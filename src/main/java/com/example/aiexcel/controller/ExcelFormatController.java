package com.example.aiexcel.controller;

import com.example.aiexcel.model.FormatOptions;
import com.example.aiexcel.service.ExcelFormatService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

/**
 * Excel格式设置控制器
 * 提供高级格式设置功能的API
 */
@RestController
@RequestMapping("/api/excel")
public class ExcelFormatController {

    @Autowired
    private ExcelFormatService excelFormatService;

    private static final Logger logger = LoggerFactory.getLogger(ExcelFormatController.class);

    /**
     * 格式化单个单元格
     */
    @PostMapping("/format-cell")
    public ResponseEntity<Map<String, Object>> formatCell(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("row") int row,
            @RequestParam("col") int col,
            @RequestBody Map<String, Object> formatOptions) {
        
        logger.info("Received request to format cell ({}, {}) in sheet {} for file: {}", 
                   row, col, sheetName, file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 创建FormatOptions对象
            FormatOptions options = excelFormatService.createFormatOptionsFromMap(formatOptions);

            // 格式化单元格
            boolean success = excelFormatService.formatCell(file, sheetName, row, col, options);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Cell formatted successfully"
                );
                logger.info("Successfully formatted cell ({}, {}) in sheet {}", row, col, sheetName);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to format cell"
                );
                logger.error("Failed to format cell ({}, {}) in sheet {}", row, col, sheetName);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error formatting cell ({}, {}) in sheet {}", row, col, sheetName, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error formatting cell: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 格式化单元格范围
     */
    @PostMapping("/format-range")
    public ResponseEntity<Map<String, Object>> formatCellRange(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("startRow") int startRow,
            @RequestParam("startCol") int startCol,
            @RequestParam("endRow") int endRow,
            @RequestParam("endCol") int endCol,
            @RequestBody Map<String, Object> formatOptions) {
        
        logger.info("Received request to format cell range ({}, {}) to ({}, {}) in sheet {} for file: {}", 
                   startRow, startCol, endRow, endCol, sheetName, file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 创建FormatOptions对象
            FormatOptions options = excelFormatService.createFormatOptionsFromMap(formatOptions);

            // 格式化单元格范围
            boolean success = excelFormatService.formatCellRange(file, sheetName, startRow, startCol, endRow, endCol, options);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Cell range formatted successfully"
                );
                logger.info("Successfully formatted cell range ({}, {}) to ({}, {}) in sheet {}", 
                           startRow, startCol, endRow, endCol, sheetName);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to format cell range"
                );
                logger.error("Failed to format cell range ({}, {}) to ({}, {}) in sheet {}", 
                            startRow, startCol, endRow, endCol, sheetName);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error formatting cell range ({}, {}) to ({}, {}) in sheet {}", 
                        startRow, startCol, endRow, endCol, sheetName, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error formatting cell range: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 格式化整行
     */
    @PostMapping("/format-row")
    public ResponseEntity<Map<String, Object>> formatRow(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("rowIndex") int rowIndex,
            @RequestBody Map<String, Object> formatOptions) {
        
        logger.info("Received request to format row {} in sheet {} for file: {}", 
                   rowIndex, sheetName, file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 创建FormatOptions对象
            FormatOptions options = excelFormatService.createFormatOptionsFromMap(formatOptions);

            // 格式化行
            boolean success = excelFormatService.formatRow(file, sheetName, rowIndex, options);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Row formatted successfully"
                );
                logger.info("Successfully formatted row {} in sheet {}", rowIndex, sheetName);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to format row"
                );
                logger.error("Failed to format row {} in sheet {}", rowIndex, sheetName);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error formatting row {} in sheet {}", rowIndex, sheetName, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error formatting row: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 格式化整列
     */
    @PostMapping("/format-column")
    public ResponseEntity<Map<String, Object>> formatColumn(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("colIndex") int colIndex,
            @RequestBody Map<String, Object> formatOptions) {
        
        logger.info("Received request to format column {} in sheet {} for file: {}", 
                   colIndex, sheetName, file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 创建FormatOptions对象
            FormatOptions options = excelFormatService.createFormatOptionsFromMap(formatOptions);

            // 格式化列
            boolean success = excelFormatService.formatColumn(file, sheetName, colIndex, options);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Column formatted successfully"
                );
                logger.info("Successfully formatted column {} in sheet {}", colIndex, sheetName);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to format column"
                );
                logger.error("Failed to format column {} in sheet {}", colIndex, sheetName);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error formatting column {} in sheet {}", colIndex, sheetName, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error formatting column: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 合并单元格并应用格式
     */
    @PostMapping("/merge-and-format")
    public ResponseEntity<Map<String, Object>> mergeAndFormatCells(
            @RequestParam("file") MultipartFile file,
            @RequestParam("sheetName") String sheetName,
            @RequestParam("startRow") int startRow,
            @RequestParam("startCol") int startCol,
            @RequestParam("endRow") int endRow,
            @RequestParam("endCol") int endCol,
            @RequestBody Map<String, Object> formatOptions) {
        
        logger.info("Received request to merge and format cells from ({}, {}) to ({}, {}) in sheet {} for file: {}", 
                   startRow, startCol, endRow, endCol, sheetName, file.getOriginalFilename());

        try {
            // 验证文件
            if (file.isEmpty()) {
                logger.error("File is empty");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "File is empty"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 创建FormatOptions对象
            FormatOptions options = excelFormatService.createFormatOptionsFromMap(formatOptions);

            // 合并并格式化单元格
            boolean success = excelFormatService.mergeAndFormatCells(file, sheetName, startRow, startCol, endRow, endCol, options);

            if (success) {
                Map<String, Object> response = Map.of(
                    "success", true,
                    "message", "Cells merged and formatted successfully"
                );
                logger.info("Successfully merged and formatted cells from ({}, {}) to ({}, {}) in sheet {}", 
                           startRow, startCol, endRow, endCol, sheetName);
                return ResponseEntity.ok(response);
            } else {
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Failed to merge and format cells"
                );
                logger.error("Failed to merge and format cells from ({}, {}) to ({}, {}) in sheet {}", 
                            startRow, startCol, endRow, endCol, sheetName);
                return ResponseEntity.badRequest().body(response);
            }
        } catch (Exception e) {
            logger.error("Error merging and formatting cells from ({}, {}) to ({}, {}) in sheet {}", 
                        startRow, startCol, endRow, endCol, sheetName, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error merging and formatting cells: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }
}