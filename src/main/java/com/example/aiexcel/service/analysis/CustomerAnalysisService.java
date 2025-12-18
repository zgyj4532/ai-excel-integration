package com.example.aiexcel.service.analysis;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Map;

public interface CustomerAnalysisService {
    Map<String, Object> performRFMAnalysis(MultipartFile file) throws IOException;
    Map<String, Object> calculateCustomerLifetimeValue(MultipartFile file) throws IOException;
    Map<String, Object> segmentCustomers(MultipartFile file) throws IOException;
    Map<String, Object> predictChurnRisk(MultipartFile file) throws IOException;
    Map<String, Object> calculateCACvsCLV(MultipartFile file) throws IOException;
    Map<String, Object> analyzeCustomerCohorts(MultipartFile file) throws IOException;
}