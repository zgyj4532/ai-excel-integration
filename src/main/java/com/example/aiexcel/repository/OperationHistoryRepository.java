package com.example.aiexcel.repository;

import com.example.aiexcel.model.OperationHistory;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.time.LocalDateTime;
import java.util.List;

/**
 * 操作历史记录仓库
 */
@Repository
public interface OperationHistoryRepository extends JpaRepository<OperationHistory, Long> {

    /**
     * 根据文件ID获取操作历史记录
     */
    List<OperationHistory> findByFileIdOrderByCreatedAtDesc(String fileId);

    /**
     * 根据文件ID和用户ID获取操作历史记录
     */
    List<OperationHistory> findByFileIdAndUserIdOrderByCreatedAtDesc(String fileId, String userId);

    /**
     * 根据文件ID和时间范围获取操作历史记录
     */
    List<OperationHistory> findByFileIdAndCreatedAtBetweenOrderByCreatedAtDesc(String fileId, LocalDateTime start, LocalDateTime end);

    /**
     * 获取指定文件的最后N条操作记录
     */
    @Query("SELECT oh FROM OperationHistory oh WHERE oh.fileId = :fileId ORDER BY oh.createdAt DESC")
    List<OperationHistory> findTopNByFileId(@Param("fileId") String fileId, org.springframework.data.domain.Pageable pageable);

    /**
     * 根据文件ID获取可撤销的操作记录
     */
    List<OperationHistory> findByFileIdAndReversibleTrueOrderByCreatedAtDesc(String fileId);

    /**
     * 获取最新的一条操作记录
     */
    OperationHistory findFirstByFileIdOrderByCreatedAtDesc(String fileId);
}