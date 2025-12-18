package com.example.aiexcel.model;

import jakarta.persistence.*;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

import java.time.LocalDateTime;

/**
 * 文件版本实体
 * 用于存储Excel文件的版本历史
 */
@Entity
@Table(name = "file_versions")
@EntityListeners(AuditingEntityListener.class)
public class FileVersion {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "file_id", nullable = false)
    private String fileId;

    @Column(name = "version_number", nullable = false)
    private Integer versionNumber;

    @Column(name = "file_name", nullable = false)
    private String fileName;

    @Column(name = "file_content", length = 100000) // 存储Base64编码的文件内容
    private String fileContent;

    @Column(name = "file_size")
    private Long fileSize;

    @Column(name = "change_description", length = 1000)
    private String changeDescription;

    @Column(name = "user_id")
    private String userId;

    @CreatedDate
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @Column(name = "is_current", nullable = false)
    private Boolean isCurrent = false;

    // Constructors
    public FileVersion() {}

    public FileVersion(String fileId, Integer versionNumber, String fileName, String fileContent, 
                       Long fileSize, String changeDescription, String userId) {
        this.fileId = fileId;
        this.versionNumber = versionNumber;
        this.fileName = fileName;
        this.fileContent = fileContent;
        this.fileSize = fileSize;
        this.changeDescription = changeDescription;
        this.userId = userId;
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

    public Integer getVersionNumber() {
        return versionNumber;
    }

    public void setVersionNumber(Integer versionNumber) {
        this.versionNumber = versionNumber;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFileContent() {
        return fileContent;
    }

    public void setFileContent(String fileContent) {
        this.fileContent = fileContent;
    }

    public Long getFileSize() {
        return fileSize;
    }

    public void setFileSize(Long fileSize) {
        this.fileSize = fileSize;
    }

    public String getChangeDescription() {
        return changeDescription;
    }

    public void setChangeDescription(String changeDescription) {
        this.changeDescription = changeDescription;
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

    public Boolean getIsCurrent() {
        return isCurrent;
    }

    public void setIsCurrent(Boolean isCurrent) {
        this.isCurrent = isCurrent;
    }
}