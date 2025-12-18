package com.example.aiexcel.model;

import jakarta.persistence.*;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

import java.time.LocalDateTime;

/**
 * 操作历史记录实体
 * 用于记录用户对Excel文件的操作历史，支持撤销/重做功能
 */
@Entity
@Table(name = "operation_history")
@EntityListeners(AuditingEntityListener.class)
public class OperationHistory {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "file_id", nullable = false)
    private String fileId;

    @Column(name = "operation_type", nullable = false)
    private String operationType;

    @Column(name = "parameters", length = 10000)
    private String parameters;

    @Column(name = "file_content_before", length = 50000)
    private String fileContentBefore;

    @Column(name = "file_content_after", length = 50000)
    private String fileContentAfter;

    @Column(name = "user_id")
    private String userId;

    @CreatedDate
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @Column(name = "reversible", nullable = false)
    private boolean reversible = true;

    // Constructors
    public OperationHistory() {}

    public OperationHistory(String fileId, String operationType, String parameters, String fileContentBefore, String fileContentAfter) {
        this.fileId = fileId;
        this.operationType = operationType;
        this.parameters = parameters;
        this.fileContentBefore = fileContentBefore;
        this.fileContentAfter = fileContentAfter;
        this.createdAt = LocalDateTime.now();
    }

    // Getters and Setters
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getFileId() {
        return fileId;
    }

    public void setFileId(String fileId) {
        this.fileId = fileId;
    }

    public String getOperationType() {
        return operationType;
    }

    public void setOperationType(String operationType) {
        this.operationType = operationType;
    }

    public String getParameters() {
        return parameters;
    }

    public void setParameters(String parameters) {
        this.parameters = parameters;
    }

    public String getFileContentBefore() {
        return fileContentBefore;
    }

    public void setFileContentBefore(String fileContentBefore) {
        this.fileContentBefore = fileContentBefore;
    }

    public String getFileContentAfter() {
        return fileContentAfter;
    }

    public void setFileContentAfter(String fileContentAfter) {
        this.fileContentAfter = fileContentAfter;
    }

    public String getUserId() {
        return userId;
    }

    public void setUserId(String userId) {
        this.userId = userId;
    }

    public LocalDateTime getCreatedAt() {
        return createdAt;
    }

    public void setCreatedAt(LocalDateTime createdAt) {
        this.createdAt = createdAt;
    }

    public boolean isReversible() {
        return reversible;
    }

    public void setReversible(boolean reversible) {
        this.reversible = reversible;
    }
}