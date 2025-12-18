package com.example.aiexcel.service.ai;

import com.example.aiexcel.dto.AiRequest;
import com.example.aiexcel.dto.AiResponse;

public interface AiService {
    AiResponse generateResponse(AiRequest request);
    AiResponse generateStreamingResponse(AiRequest request);
    boolean testConnection();
}