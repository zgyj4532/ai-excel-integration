package com.example.aiexcel.service;

import com.example.aiexcel.dto.UndoRedoResult;
import com.example.aiexcel.model.OperationHistory;
import com.example.aiexcel.repository.OperationHistoryRepository;
import com.example.aiexcel.service.excel.ExcelService;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import java.util.Base64;
import java.util.List;

/**
 * 操作历史记录服务
 * 提供操作记录、撤销、重做等功能
 */
@Service
public class OperationHistoryService {

    @Autowired
    private OperationHistoryRepository operationHistoryRepository;

    @Autowired
    private ExcelService excelService;

    private static final Logger logger = LoggerFactory.getLogger(OperationHistoryService.class);

    /**
     * 记录操作历史
     */
    public OperationHistory recordOperation(String fileId, String operationType, String parameters, 
                                           MultipartFile fileBefore, MultipartFile fileAfter) {
        try {
            String contentBefore = fileBefore != null ? encodeFileContent(fileBefore) : null;
            String contentAfter = fileAfter != null ? encodeFileContent(fileAfter) : null;

            OperationHistory history = new OperationHistory(fileId, operationType, parameters, contentBefore, contentAfter);
            history = operationHistoryRepository.save(history);

            logger.info("Recorded operation: {} for file: {}", operationType, fileId);
            return history;
        } catch (Exception e) {
            logger.error("Error recording operation for file: {}", fileId, e);
            return null;
        }
    }

    /**
     * 记录操作历史，传入Workbook对象
     */
    public OperationHistory recordOperationWithWorkbooks(String fileId, String operationType, String parameters, 
                                                        Workbook workbookBefore, Workbook workbookAfter) {
        try {
            String contentBefore = workbookBefore != null ? encodeWorkbookContent(workbookBefore) : null;
            String contentAfter = workbookAfter != null ? encodeWorkbookContent(workbookAfter) : null;

            OperationHistory history = new OperationHistory(fileId, operationType, parameters, contentBefore, contentAfter);
            history = operationHistoryRepository.save(history);

            logger.info("Recorded operation: {} for file: {}", operationType, fileId);
            return history;
        } catch (Exception e) {
            logger.error("Error recording operation for file: {}", fileId, e);
            return null;
        }
    }

    /**
     * 获取文件的操作历史列表
     */
    public List<OperationHistory> getOperationHistory(String fileId) {
        try {
            List<OperationHistory> history = operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId);
            logger.debug("Retrieved {} operation history records for file: {}", history.size(), fileId);
            return history;
        } catch (Exception e) {
            logger.error("Error retrieving operation history for file: {}", fileId, e);
            return List.of();
        }
    }

    /**
     * 获取文件的可撤销操作历史列表
     */
    public List<OperationHistory> getReversibleOperations(String fileId) {
        try {
            List<OperationHistory> history = operationHistoryRepository.findByFileIdAndReversibleTrueOrderByCreatedAtDesc(fileId);
            logger.debug("Retrieved {} reversible operation history records for file: {}", history.size(), fileId);
            return history;
        } catch (Exception e) {
            logger.error("Error retrieving reversible operations for file: {}", fileId, e);
            return List.of();
        }
    }

    /**
     * 获取最近的N条操作历史
     */
    public List<OperationHistory> getRecentOperations(String fileId, int count) {
        try {
            //Pageable pageable = PageRequest.of(0, count, Sort.by(Sort.Direction.DESC, "createdAt"));
            // 由于Spring Data JPA不直接支持top N查询的Pageable，我们使用另一种方式
            List<OperationHistory> allHistory = operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId);
            if (allHistory.size() <= count) {
                return allHistory;
            }
            return allHistory.subList(0, count);
        } catch (Exception e) {
            logger.error("Error retrieving recent operations for file: {}", fileId, e);
            return List.of();
        }
    }

    /**
     * 撤销上一次操作，恢复文件到操作前的状态
     */
    public UndoRedoResult undoLastOperation(String fileId) {
        try {
            OperationHistory lastOperation = operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc(fileId);
            if (lastOperation == null) {
                logger.warn("No operation found to undo for file: {}", fileId);
                return new UndoRedoResult(false, "No operation found to undo");
            }

            if (!lastOperation.isReversible()) {
                logger.warn("Last operation is not reversible for file: {}", fileId);
                return new UndoRedoResult(false, "Last operation is not reversible");
            }

            if (lastOperation.isUndone()) {
                logger.warn("Last operation is already undone for file: {}", fileId);
                return new UndoRedoResult(false, "Last operation is already undone");
            }

            // 恢复到操作前的状态
            String contentBefore = lastOperation.getFileContentBefore();
            if (contentBefore == null || contentBefore.isEmpty()) {
                logger.warn("No previous content to restore for operation: {} on file: {}", lastOperation.getId(), fileId);
                return new UndoRedoResult(false, "No previous content to restore");
            }

            // 从Base64内容恢复工作簿并保存
            Workbook restoredWorkbook = createWorkbookFromContent(contentBefore);
            if (restoredWorkbook == null) {
                logger.error("Failed to create workbook from stored content for file: {}", fileId);
                return new UndoRedoResult(false, "Failed to restore workbook from stored content");
            }

            String restoredFileName = "restored_" + fileId + ".xlsx";
            excelService.saveWorkbook(restoredWorkbook, restoredFileName);

            // 标记此操作为已撤销
            lastOperation.setUndone(true);
            lastOperation.setReversible(false);
            operationHistoryRepository.save(lastOperation);

            logger.info("Undid operation: {} for file: {}, content restored", lastOperation.getOperationType(), fileId);
            return new UndoRedoResult(true, "Successfully undid operation: " + lastOperation.getOperationType(),
                    contentBefore, lastOperation.getOperationType());
        } catch (Exception e) {
            logger.error("Error undoing last operation for file: {}", fileId, e);
            return new UndoRedoResult(false, "Error undoing: " + e.getMessage());
        }
    }

    /**
     * 重做上次被撤销的操作，恢复文件到操作后的状态
     */
    public UndoRedoResult redoLastUndoneOperation(String fileId) {
        try {
            OperationHistory undoneOperation = operationHistoryRepository.findFirstByFileIdAndUndoneTrueOrderByCreatedAtDesc(fileId);
            if (undoneOperation == null) {
                logger.warn("No undone operation found to redo for file: {}", fileId);
                return new UndoRedoResult(false, "No undone operation found to redo");
            }

            String contentAfter = undoneOperation.getFileContentAfter();
            if (contentAfter == null || contentAfter.isEmpty()) {
                logger.warn("No 'after' content to restore for redo on file: {}", fileId);
                return new UndoRedoResult(false, "No content to restore for redo");
            }

            // 从Base64内容恢复工作簿并保存
            Workbook redoneWorkbook = createWorkbookFromContent(contentAfter);
            if (redoneWorkbook == null) {
                logger.error("Failed to create workbook from stored after-content for file: {}", fileId);
                return new UndoRedoResult(false, "Failed to restore workbook from stored content");
            }

            String redoneFileName = "redone_" + fileId + ".xlsx";
            excelService.saveWorkbook(redoneWorkbook, redoneFileName);

            // 取消撤销标记，恢复可撤销状态
            undoneOperation.setUndone(false);
            undoneOperation.setReversible(true);
            operationHistoryRepository.save(undoneOperation);

            logger.info("Redid operation: {} for file: {}", undoneOperation.getOperationType(), fileId);
            return new UndoRedoResult(true, "Successfully redid operation: " + undoneOperation.getOperationType(),
                    contentAfter, undoneOperation.getOperationType());
        } catch (Exception e) {
            logger.error("Error redoing last undone operation for file: {}", fileId, e);
            return new UndoRedoResult(false, "Error redoing: " + e.getMessage());
        }
    }

    /**
     * 获取最后一次操作
     */
    public OperationHistory getLastOperation(String fileId) {
        try {
            return operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc(fileId);
        } catch (Exception e) {
            logger.error("Error retrieving last operation for file: {}", fileId, e);
            return null;
        }
    }

    /**
     * 清除文件的所有操作历史
     */
    public void clearOperationHistory(String fileId) {
        try {
            List<OperationHistory> histories = operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId);
            operationHistoryRepository.deleteAll(histories);
            logger.info("Cleared operation history for file: {}", fileId);
        } catch (Exception e) {
            logger.error("Error clearing operation history for file: {}", fileId, e);
        }
    }

    /**
     * 将MultipartFile编码为Base64字符串
     */
    private String encodeFileContent(MultipartFile file) throws IOException {
        byte[] fileBytes = file.getBytes();
        return Base64.getEncoder().encodeToString(fileBytes);
    }

    /**
     * 将Workbook编码为Base64字符串
     */
    private String encodeWorkbookContent(Workbook workbook) throws IOException {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            workbook.write(baos);
            byte[] workbookBytes = baos.toByteArray();
            return Base64.getEncoder().encodeToString(workbookBytes);
        }
    }

    /**
     * 将Base64字符串解码为字节数组
     */
    private byte[] decodeContent(String encodedContent) {
        if (encodedContent == null || encodedContent.isEmpty()) {
            return new byte[0];
        }
        return Base64.getDecoder().decode(encodedContent);
    }

    /**
     * 从编码的内容创建Workbook
     */
    private Workbook createWorkbookFromContent(String encodedContent) throws IOException {
        byte[] content = decodeContent(encodedContent);
        if (content.length == 0) {
            return null;
        }
        return excelService.loadWorkbook(new java.io.ByteArrayInputStream(content));
    }
}