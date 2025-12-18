package com.example.aiexcel.repository;

import com.example.aiexcel.model.WorkspaceFile;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * 工作区文件仓库
 */
@Repository
public interface WorkspaceFileRepository extends JpaRepository<WorkspaceFile, Long> {

    /**
     * 根据工作区ID获取文件列表
     */
    List<WorkspaceFile> findByWorkspaceId(Long workspaceId);

    /**
     * 根据工作区ID和上传者获取文件列表
     */
    List<WorkspaceFile> findByWorkspaceIdAndUploadedBy(Long workspaceId, String uploadedBy);

    /**
     * 根据上传者获取文件列表
     */
    List<WorkspaceFile> findByUploadedBy(String uploadedBy);

    /**
     * 根据文件类型获取文件列表
     */
    List<WorkspaceFile> findByFileType(String fileType);

    /**
     * 根据工作区ID和文件类型获取文件列表
     */
    List<WorkspaceFile> findByWorkspaceIdAndFileType(Long workspaceId, String fileType);

    /**
     * 根据工作区ID和是否收藏获取文件列表
     */
    List<WorkspaceFile> findByWorkspaceIdAndIsFavoriteTrue(Long workspaceId);

    /**
     * 根据工作区ID和是否归档获取文件列表
     */
    List<WorkspaceFile> findByWorkspaceIdAndIsArchivedTrue(Long workspaceId);

    /**
     * 根据文件名搜索工作区文件
     */
    List<WorkspaceFile> findByWorkspaceIdAndFileNameContainingIgnoreCase(Long workspaceId, String fileName);

    /**
     * 获取最近的文件（按创建时间排序）
     */
    List<WorkspaceFile> findTop10ByWorkspaceIdOrderByCreatedAtDesc(Long workspaceId);

    /**
     * 获取最大的文件（按文件大小排序）
     */
    @Query("SELECT wf FROM WorkspaceFile wf WHERE wf.workspaceId = :workspaceId ORDER BY wf.fileSize DESC")
    List<WorkspaceFile> findByWorkspaceIdOrderByFileSizeDesc(@Param("workspaceId") Long workspaceId);

    /**
     * 获取工作区中文件总数
     */
    int countByWorkspaceId(Long workspaceId);
}