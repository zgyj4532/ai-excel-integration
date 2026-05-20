package com.example.aiexcel.unit.service;

import com.example.aiexcel.dto.UndoRedoResult;
import com.example.aiexcel.model.OperationHistory;
import com.example.aiexcel.repository.OperationHistoryRepository;
import com.example.aiexcel.service.OperationHistoryService;
import com.example.aiexcel.service.excel.ExcelService;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class OperationHistoryServiceTest {

    @Mock
    private OperationHistoryRepository operationHistoryRepository;

    @Mock
    private ExcelService excelService;

    @InjectMocks
    private OperationHistoryService operationHistoryService;

    @Test
    void testRecordOperationWithWorkbooks_Success() {
        String fileId = "test-file-1";
        String operationType = "SET_CELL";
        String parameters = "A1=test";

        Workbook mockWorkbook = mock(Workbook.class);
        when(operationHistoryRepository.save(any(OperationHistory.class)))
                .thenAnswer(invocation -> {
                    OperationHistory h = invocation.getArgument(0);
                    h.setId(1L);
                    return h;
                });

        OperationHistory result = operationHistoryService.recordOperationWithWorkbooks(
                fileId, operationType, parameters, null, mockWorkbook);

        assertNotNull(result);
        assertEquals(fileId, result.getFileId());
        assertEquals(operationType, result.getOperationType());
        verify(operationHistoryRepository, times(1)).save(any(OperationHistory.class));
    }

    @Test
    void testRecordOperationWithWorkbooks_NullWorkbooks() {
        when(operationHistoryRepository.save(any(OperationHistory.class)))
                .thenAnswer(invocation -> {
                    OperationHistory h = invocation.getArgument(0);
                    h.setId(2L);
                    return h;
                });

        OperationHistory result = operationHistoryService.recordOperationWithWorkbooks(
                "file-1", "INSERT_ROW", "3", null, null);

        assertNotNull(result);
        assertNull(result.getFileContentBefore());
        assertNull(result.getFileContentAfter());
    }

    @Test
    void testGetOperationHistory_ReturnsList() {
        String fileId = "test-file-1";
        OperationHistory h1 = new OperationHistory(fileId, "SET_CELL", "A1=1", null, null);
        OperationHistory h2 = new OperationHistory(fileId, "INSERT_ROW", "2", null, null);
        when(operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId))
                .thenReturn(List.of(h2, h1));

        List<OperationHistory> result = operationHistoryService.getOperationHistory(fileId);

        assertEquals(2, result.size());
        verify(operationHistoryRepository, times(1)).findByFileIdOrderByCreatedAtDesc(fileId);
    }

    @Test
    void testGetOperationHistory_EmptyResult() {
        when(operationHistoryRepository.findByFileIdOrderByCreatedAtDesc("empty"))
                .thenReturn(List.of());

        List<OperationHistory> result = operationHistoryService.getOperationHistory("empty");

        assertTrue(result.isEmpty());
    }

    @Test
    void testGetReversibleOperations_ReturnsList() {
        String fileId = "test-file-1";
        OperationHistory h = new OperationHistory(fileId, "SET_CELL", "A1=1", null, null);
        h.setReversible(true);
        when(operationHistoryRepository.findByFileIdAndReversibleTrueOrderByCreatedAtDesc(fileId))
                .thenReturn(List.of(h));

        List<OperationHistory> result = operationHistoryService.getReversibleOperations(fileId);

        assertEquals(1, result.size());
        assertTrue(result.get(0).isReversible());
    }

    @Test
    void testGetRecentOperations_LessThanCount() {
        String fileId = "test-file-1";
        OperationHistory h = new OperationHistory(fileId, "SET_CELL", "A1=1", null, null);
        when(operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId))
                .thenReturn(List.of(h));

        List<OperationHistory> result = operationHistoryService.getRecentOperations(fileId, 5);

        assertEquals(1, result.size());
    }

    @Test
    void testGetRecentOperations_MoreThanCount() {
        String fileId = "test-file-1";
        List<OperationHistory> many = List.of(
                new OperationHistory(fileId, "OP1", "", null, null),
                new OperationHistory(fileId, "OP2", "", null, null),
                new OperationHistory(fileId, "OP3", "", null, null)
        );
        when(operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId))
                .thenReturn(many);

        List<OperationHistory> result = operationHistoryService.getRecentOperations(fileId, 2);

        assertEquals(2, result.size());
    }

    @Test
    void testUndoLastOperation_NoHistory() {
        when(operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc("empty"))
                .thenReturn(null);

        UndoRedoResult result = operationHistoryService.undoLastOperation("empty");

        assertFalse(result.isSuccess());
    }

    @Test
    void testUndoLastOperation_NotReversible() {
        OperationHistory h = new OperationHistory("f1", "SET_CELL", "A1=1", "before", "after");
        h.setReversible(false);
        when(operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc("f1"))
                .thenReturn(h);

        UndoRedoResult result = operationHistoryService.undoLastOperation("f1");

        assertFalse(result.isSuccess());
    }

    @Test
    void testUndoLastOperation_NoPreviousContent() {
        OperationHistory h = new OperationHistory("f1", "SET_CELL", "A1=1", null, "after");
        h.setReversible(true);
        when(operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc("f1"))
                .thenReturn(h);

        UndoRedoResult result = operationHistoryService.undoLastOperation("f1");

        assertFalse(result.isSuccess());
    }

    @Test
    void testRedoLastUndoneOperation_NoUndoneOp() {
        when(operationHistoryRepository.findFirstByFileIdAndUndoneTrueOrderByCreatedAtDesc("f1"))
                .thenReturn(null);

        UndoRedoResult result = operationHistoryService.redoLastUndoneOperation("f1");
        assertFalse(result.isSuccess());
    }

    @Test
    void testGetLastOperation_ReturnsOperation() {
        OperationHistory h = new OperationHistory("f1", "SET_CELL", "A1=1", null, null);
        when(operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc("f1"))
                .thenReturn(h);

        OperationHistory result = operationHistoryService.getLastOperation("f1");

        assertNotNull(result);
        assertEquals("SET_CELL", result.getOperationType());
    }

    @Test
    void testGetLastOperation_Null() {
        when(operationHistoryRepository.findFirstByFileIdOrderByCreatedAtDesc("empty"))
                .thenReturn(null);

        OperationHistory result = operationHistoryService.getLastOperation("empty");

        assertNull(result);
    }

    @Test
    void testClearOperationHistory() {
        String fileId = "f1";
        OperationHistory h = new OperationHistory(fileId, "SET_CELL", "A1=1", null, null);
        when(operationHistoryRepository.findByFileIdOrderByCreatedAtDesc(fileId))
                .thenReturn(List.of(h));

        operationHistoryService.clearOperationHistory(fileId);

        verify(operationHistoryRepository, times(1)).deleteAll(anyList());
    }
}
