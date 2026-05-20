package com.example.aiexcel.dto;

/**
 * 撤销/重做操作结果
 */
public class UndoRedoResult {
    private boolean success;
    private String message;
    private String restoredContent;
    private String operationType;

    public UndoRedoResult(boolean success, String message) {
        this.success = success;
        this.message = message;
    }

    public UndoRedoResult(boolean success, String message, String restoredContent, String operationType) {
        this.success = success;
        this.message = message;
        this.restoredContent = restoredContent;
        this.operationType = operationType;
    }

    public boolean isSuccess() { return success; }
    public void setSuccess(boolean success) { this.success = success; }

    public String getMessage() { return message; }
    public void setMessage(String message) { this.message = message; }

    public String getRestoredContent() { return restoredContent; }
    public void setRestoredContent(String restoredContent) { this.restoredContent = restoredContent; }

    public String getOperationType() { return operationType; }
    public void setOperationType(String operationType) { this.operationType = operationType; }
}
