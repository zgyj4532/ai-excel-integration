package com.example.aiexcel.service.analysis;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Map;

public interface FinancialAnalysisService {
    Map<String, Object> analyzeFinancialStatements(MultipartFile file, String analysisType) throws IOException;
    Map<String, Object> calculateFinancialRatios(MultipartFile file) throws IOException;
    Map<String, Object> analyzeProfitability(MultipartFile file) throws IOException;
    Map<String, Object> analyzeCashFlow(MultipartFile file) throws IOException;
    Map<String, Object> compareBudgetVsActual(MultipartFile file) throws IOException;
}