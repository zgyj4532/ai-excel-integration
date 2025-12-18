package com.example.aiexcel.model;

import jakarta.persistence.*;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.annotation.LastModifiedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

import java.time.LocalDateTime;

/**
 * 工作区中的文件项实体
 * 表示存储在工作区中的具体文件
 */
@Entity
@Table(name = "workspace_files")
@EntityListeners(AuditingEntityListener.class)
public class WorkspaceFile {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "file_name", nullable = false)
    private String fileName;

    @Column(name = "file_path", length = 2000)
    private String filePath;

    @Column(name = "file_size")
    private Long fileSize;

    @Column(name = "file_type")
    private String fileType;

    @Column(name = "workspace_id", nullable = false)
    private Long workspaceId;

    @Column(name = "uploaded_by", nullable = false)
    private String uploadedBy;

    @CreatedDate
    @Column(name = "created_at", nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @LastModifiedDate
    @Column(name = "updated_at")
    private LocalDateTime updatedAt;

    @Column(name = "description", length = 1000)
    private String description;

    @Column(name = "tags", length = 500)
    private String tags;

    @Column(name = "version")
    private String version;

    @Column(name = "is_favorite")
    private Boolean isFavorite = false;

    @Column(name = "is_archived")
    private Boolean isArchived = false;

    // Constructors
    public WorkspaceFile() {}

    public WorkspaceFile(String fileName, String filePath, Long fileSize, String fileType, 
                        Long workspaceId, String uploadedBy) {
        this.fileName = fileName;
        this.filePath = filePath;
        this.fileSize = fileSize;
        this.fileType = fileType;
        this.workspaceId = workspaceId;
        this.uploadedBy = uploadedBy;
        this.createdAt = LocalDateTime.now();
        this.updatedAt = LocalDateTime.now();
    }

    public WorkspaceFile(String fileName, String filePath, Long fileSize, String fileType, 
                        Long workspaceId, String uploadedBy, String description) {
        this.fileName = fileName;
        this.filePath = filePath;
        this.fileSize = fileSize;
        this.fileType = fileType;
        this.workspaceId = workspaceId;
        this.uploadedBy = uploadedBy;
        this.description = description;
        this.createdAt = LocalDateTime.now();
        this.updatedAt = LocalDateTime.now();
    }

    // Getters and Setters
    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public Long getFileSize() {
        return fileSize;
    }

    public void setFileSize(Long fileSize) {
        this.fileSize = fileSize;
    }

    public String getFileType() {
        return fileType;
    }

    public void setFileType(String fileType) {
        this.fileType = fileType;
    }

    public Long getWorkspaceId() {
        return workspaceId;
    }

    public void setWorkspaceId(Long workspaceId) {
        this.workspaceId = workspaceId;
    }

    public String getUploadedBy() {
        return uploadedBy;
    }

    public void setUploadedBy(String uploadedBy) {
        this.uploadedBy = uploadedBy;
    }

    public LocalDateTime getCreatedAt() {
        return createdAt;
    }

    public void setCreatedAt(LocalDateTime createdAt) {
        this.createdAt = createdAt;
    }

    public LocalDateTime getUpdatedAt() {
        return updatedAt;
    }

    public void setUpdatedAt(LocalDateTime updatedAt) {
        this.updatedAt = updatedAt;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getTags() {
        return tags;
    }

    public void setTags(String tags) {
        this.tags = tags;
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public Boolean getIsFavorite() {
        return isFavorite;
    }

    public void setIsFavorite(Boolean isFavorite) {
        this.isFavorite = isFavorite;
    }

    public Boolean getIsArchived() {
        return isArchived;
    }

    public void setIsArchived(Boolean isArchived) {
        this.isArchived = isArchived;
    }
}