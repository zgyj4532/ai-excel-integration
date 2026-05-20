package com.example.aiexcel.unit.controller;

import com.example.aiexcel.controller.OperationHistoryController;
import com.example.aiexcel.dto.UndoRedoResult;
import com.example.aiexcel.model.OperationHistory;
import com.example.aiexcel.service.OperationHistoryService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.http.ResponseEntity;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class OperationHistoryControllerTest {

    @Mock
    private OperationHistoryService operationHistoryService;

    @InjectMocks
    private OperationHistoryController controller;

    @Test
    void testGetOperationHistory_Success() {
        OperationHistory h = new OperationHistory("file1", "SET_CELL", "A1=1", null, null);
        when(operationHistoryService.getOperationHistory("file1")).thenReturn(List.of(h));

        ResponseEntity<Map<String, Object>> response = controller.getOperationHistory("file1");

        assertEquals(200, response.getStatusCode().value());
        assertTrue((Boolean) response.getBody().get("success"));
        assertEquals(1, response.getBody().get("count"));
    }

    @Test
    void testGetReversibleOperations_Success() {
        OperationHistory h = new OperationHistory("file1", "SET_CELL", "A1=1", null, null);
        h.setReversible(true);
        when(operationHistoryService.getReversibleOperations("file1")).thenReturn(List.of(h));

        ResponseEntity<Map<String, Object>> response = controller.getReversibleOperations("file1");

        assertEquals(200, response.getStatusCode().value());
        assertTrue((Boolean) response.getBody().get("success"));
    }

    @Test
    void testGetRecentOperations_Success() {
        when(operationHistoryService.getRecentOperations("file1", 5)).thenReturn(List.of());

        ResponseEntity<Map<String, Object>> response = controller.getRecentOperations("file1", 5);

        assertEquals(200, response.getStatusCode().value());
        assertEquals(0, response.getBody().get("count"));
    }

    @Test
    void testUndoLastOperation_Success() {
        UndoRedoResult undoResult = new UndoRedoResult(true, "Successfully undid operation", "content", "SET_CELL");
        when(operationHistoryService.undoLastOperation("file1")).thenReturn(undoResult);

        ResponseEntity<Map<String, Object>> response = controller.undoLastOperation("file1");

        assertEquals(200, response.getStatusCode().value());
        assertTrue((Boolean) response.getBody().get("success"));
    }

    @Test
    void testUndoLastOperation_Failure() {
        UndoRedoResult undoResult = new UndoRedoResult(false, "No operation found");
        when(operationHistoryService.undoLastOperation("file1")).thenReturn(undoResult);

        ResponseEntity<Map<String, Object>> response = controller.undoLastOperation("file1");

        assertEquals(200, response.getStatusCode().value());
        assertFalse((Boolean) response.getBody().get("success"));
    }

    @Test
    void testRedoLastUndoneOperation_Failure() {
        UndoRedoResult redoResult = new UndoRedoResult(false, "No undone operation");
        when(operationHistoryService.redoLastUndoneOperation("file1")).thenReturn(redoResult);

        ResponseEntity<Map<String, Object>> response = controller.redoLastUndoneOperation("file1");

        assertEquals(200, response.getStatusCode().value());
        assertFalse((Boolean) response.getBody().get("success"));
    }

    @Test
    void testClearOperationHistory() {
        doNothing().when(operationHistoryService).clearOperationHistory("file1");

        ResponseEntity<Map<String, Object>> response = controller.clearOperationHistory("file1");

        assertEquals(200, response.getStatusCode().value());
        assertTrue((Boolean) response.getBody().get("success"));
    }
}
