package com.example.aiexcel.controller;

import com.example.aiexcel.model.OperationHistory;
import com.example.aiexcel.service.OperationHistoryService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

/**
 * 操作历史记录控制器
 * 提供操作历史记录、查看、撤销等功能的API
 */
@RestController
@RequestMapping("/api/history")
public class OperationHistoryController {

    @Autowired
    private OperationHistoryService operationHistoryService;

    private static final Logger logger = LoggerFactory.getLogger(OperationHistoryController.class);

    /**
     * 获取文件的操作历史记录
     */
    @GetMapping("/file/{fileId}")
    public ResponseEntity<Map<String, Object>> getOperationHistory(@PathVariable String fileId) {
        logger.info("Received request to get operation history for file: {}", fileId);

        try {
            List<OperationHistory> history = operationHistoryService.getOperationHistory(fileId);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", history,
                "count", history.size()
            );

            logger.info("Successfully returned {} operation history records for file: {}", history.size(), fileId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving operation history for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving operation history: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取文件的可撤销操作历史记录
     */
    @GetMapping("/reversible/{fileId}")
    public ResponseEntity<Map<String, Object>> getReversibleOperations(@PathVariable String fileId) {
        logger.info("Received request to get reversible operations for file: {}", fileId);

        try {
            List<OperationHistory> history = operationHistoryService.getReversibleOperations(fileId);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", history,
                "count", history.size()
            );

            logger.info("Successfully returned {} reversible operations for file: {}", history.size(), fileId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving reversible operations for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving reversible operations: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 获取最近的N条操作历史记录
     */
    @GetMapping("/recent/{fileId}/{count}")
    public ResponseEntity<Map<String, Object>> getRecentOperations(@PathVariable String fileId, @PathVariable int count) {
        logger.info("Received request to get {} recent operations for file: {}", count, fileId);

        try {
            List<OperationHistory> history = operationHistoryService.getRecentOperations(fileId, count);

            Map<String, Object> response = Map.of(
                "success", true,
                "data", history,
                "count", history.size()
            );

            logger.info("Successfully returned {} recent operations for file: {}", history.size(), fileId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error retrieving recent operations for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error retrieving recent operations: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 撤销最后一次操作
     */
    @PostMapping("/undo/{fileId}")
    public ResponseEntity<Map<String, Object>> undoLastOperation(@PathVariable String fileId) {
        logger.info("Received request to undo last operation for file: {}", fileId);

        try {
            boolean success = operationHistoryService.undoLastOperation(fileId);

            Map<String, Object> response;
            if (success) {
                response = Map.of(
                    "success", true,
                    "message", "Successfully undid last operation"
                );
                logger.info("Successfully undid last operation for file: {}", fileId);
            } else {
                response = Map.of(
                    "success", false,
                    "error", "Could not undo last operation"
                );
                logger.warn("Could not undo last operation for file: {}", fileId);
            }

            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error undoing last operation for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error undoing last operation: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 重做最后一次撤销的操作
     */
    @PostMapping("/redo/{fileId}")
    public ResponseEntity<Map<String, Object>> redoLastUndoneOperation(@PathVariable String fileId) {
        logger.info("Received request to redo last operation for file: {}", fileId);

        try {
            boolean success = operationHistoryService.redoLastUndoneOperation(fileId);

            Map<String, Object> response;
            if (success) {
                response = Map.of(
                    "success", true,
                    "message", "Successfully redid last operation"
                );
                logger.info("Successfully redid last operation for file: {}", fileId);
            } else {
                response = Map.of(
                    "success", false,
                    "error", "Could not redo last operation"
                );
                logger.warn("Could not redo last operation for file: {}", fileId);
            }

            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error redoing last operation for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error redoing last operation: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }

    /**
     * 清除文件的所有操作历史记录
     */
    @DeleteMapping("/clear/{fileId}")
    public ResponseEntity<Map<String, Object>> clearOperationHistory(@PathVariable String fileId) {
        logger.info("Received request to clear operation history for file: {}", fileId);

        try {
            operationHistoryService.clearOperationHistory(fileId);

            Map<String, Object> response = Map.of(
                "success", true,
                "message", "Successfully cleared operation history"
            );

            logger.info("Successfully cleared operation history for file: {}", fileId);
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            logger.error("Error clearing operation history for file: {}", fileId, e);
            Map<String, Object> response = Map.of(
                "success", false,
                "error", "Error clearing operation history: " + e.getMessage()
            );
            return ResponseEntity.badRequest().body(response);
        }
    }
}