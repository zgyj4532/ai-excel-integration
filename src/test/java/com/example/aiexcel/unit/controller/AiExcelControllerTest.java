package com.example.aiexcel.unit.controller;

import com.example.aiexcel.controller.AiExcelController;
import com.example.aiexcel.service.AiAdvancedOperationsService;
import com.example.aiexcel.service.AiExcelIntegrationService;
import com.example.aiexcel.service.ai.AiService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.http.ResponseEntity;
import org.springframework.mock.web.MockMultipartFile;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class AiExcelControllerTest {

    @Mock
    private AiExcelIntegrationService aiExcelIntegrationService;

    @Mock
    private AiAdvancedOperationsService aiAdvancedOperationsService;

    @Mock
    private AiService aiService;

    @InjectMocks
    private AiExcelController controller;

    @Test
    void testHealthCheck() {
        ResponseEntity<Map<String, Object>> response = controller.healthCheck();

        assertEquals(200, response.getStatusCode().value());
        assertEquals("UP", response.getBody().get("status"));
        assertEquals("AI Excel Integration", response.getBody().get("service"));
    }

    @Test
    void testUploadExcel_Success() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "test.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "test content".getBytes());

        ResponseEntity<Map<String, Object>> response = controller.uploadExcel(file);

        assertEquals(200, response.getStatusCode().value());
        assertTrue((Boolean) response.getBody().get("success"));
    }
}
