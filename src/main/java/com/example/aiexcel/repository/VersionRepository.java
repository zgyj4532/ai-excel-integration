package com.example.aiexcel.repository;

import com.example.aiexcel.model.FileVersion;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * 文件版本仓库
 */
@Repository
public interface VersionRepository extends JpaRepository<FileVersion, Long> {

    /**
     * 根据文件ID获取所有版本
     */
    List<FileVersion> findByFileIdOrderByVersionNumberDesc(String fileId);

    /**
     * 根据文件ID和用户ID获取版本
     */
    List<FileVersion> findByFileIdAndUserIdOrderByVersionNumberDesc(String fileId, String userId);

    /**
     * 根据文件ID获取最新版本
     */
    FileVersion findTopByFileIdOrderByVersionNumberDesc(String fileId);

    /**
     * 根据文件ID和版本号获取版本
     */
    FileVersion findByFileIdAndVersionNumber(String fileId, Integer versionNumber);

    /**
     * 获取当前活跃版本
     */
    FileVersion findByFileIdAndIsCurrentTrue(String fileId);

    /**
     * 获取最新版本号
     */
    Integer findMaxVersionNumberByFileId(String fileId);
}