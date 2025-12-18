package com.example.aiexcel.controller;

import com.example.aiexcel.service.ExcelPreviewService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Map;

/**
 * Excel预览控制器
 * 提供Excel表格的可视化预览功能
 */
@RestController
@RequestMapping("/api/excel")
public class ExcelPreviewController {

    @Autowired
    private ExcelPreviewService excelPreviewService;

    private static final Logger logger = LoggerFactory.getLogger(ExcelPreviewController.class);

    /**
     * 获取Excel文件的预览数据
     */
    @GetMapping("/preview")
    public ResponseEntity<Map<String, Object>> previewExcel(@RequestParam("file") MultipartFile file) {
        logger.info("Received request to preview Excel file: {}", file.getOriginalFilename());

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

            // 获取预览数据
            Map<String, Object> previewData = excelPreviewService.getExcelPreviewData(file);

            logger.info("Successfully returned preview data for file: {}", file.getOriginalFilename());
            return ResponseEntity.ok(previewData);
        } catch (IOException e) {
            logger.error("IO error while previewing Excel file: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error reading Excel file: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        } catch (Exception e) {
            logger.error("Error while previewing Excel file: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error processing Excel file: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取单元格格式信息
     */
    @GetMapping("/cell-format")
    public ResponseEntity<Map<String, Object>> getCellFormat(
            @RequestParam("file") MultipartFile file,
            @RequestParam("row") int row,
            @RequestParam("col") int col) {
        logger.info("Received request to get cell format for file: {}, row: {}, col: {}", 
                   file.getOriginalFilename(), row, col);

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

            // 获取单元格格式信息
            Map<String, Object> formatInfo = excelPreviewService.getCellFormat(file, row, col);
            formatInfo.put("success", true);

            logger.info("Successfully returned format info for cell({}, {})", row, col);
            return ResponseEntity.ok(formatInfo);
        } catch (IOException e) {
            logger.error("IO error while getting cell format: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error reading Excel file: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        } catch (Exception e) {
            logger.error("Error while getting cell format: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error processing Excel file: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取多个单元格的格式信息（批量获取）
     */
    @PostMapping("/bulk-cell-format")
    public ResponseEntity<Map<String, Object>> getBulkCellFormat(
            @RequestParam("file") MultipartFile file,
            @RequestBody Map<String, Object> requestBody) {
        logger.info("Received request to get bulk cell format for file: {}", file.getOriginalFilename());

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

            // 获取单元格范围信息
            @SuppressWarnings("unchecked")
            Map<String, Integer> range = (Map<String, Integer>) requestBody.get("range");
            if (range == null) {
                logger.error("Range information is null");
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Range information is required"
                );
                return ResponseEntity.badRequest().body(response);
            }

            int startRow = range.getOrDefault("startRow", 0);
            int startCol = range.getOrDefault("startCol", 0);
            int endRow = range.getOrDefault("endRow", startRow);
            int endCol = range.getOrDefault("endCol", startCol);

            // 验证范围参数
            if (startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol) {
                logger.error("Invalid range parameters: startRow={}, startCol={}, endRow={}, endCol={}",
                            startRow, startCol, endRow, endCol);
                Map<String, Object> response = Map.of(
                    "success", false,
                    "error", "Invalid range parameters"
                );
                return ResponseEntity.badRequest().body(response);
            }

            // 调用服务获取批量格式信息
            Map<String, Object> formatData = excelPreviewService.getBulkCellFormat(
                file, startRow, startCol, endRow, endCol);

            formatData.put("success", true);

            logger.info("Successfully returned bulk cell format data for range ({}, {}) to ({}, {})",
                       startRow, startCol, endRow, endCol);
            return ResponseEntity.ok(formatData);
        } catch (IOException e) {
            logger.error("IO error while getting bulk cell format: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error reading Excel file: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        } catch (Exception e) {
            logger.error("Error while getting bulk cell format: {}", e.getMessage(), e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error processing request: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }
}