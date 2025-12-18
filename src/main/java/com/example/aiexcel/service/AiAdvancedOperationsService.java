package com.example.aiexcel.service;

import com.example.aiexcel.dto.AiRequest;
import com.example.aiexcel.dto.AiResponse;
import com.example.aiexcel.service.ai.AiService;
import com.example.aiexcel.service.excel.ExcelService;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * AI高级操作服务
 * 提供更丰富的AI驱动Excel操作功能
 */
@Service
public class AiAdvancedOperationsService {

    private static final Logger logger = LoggerFactory.getLogger(AiAdvancedOperationsService.class);

    @Autowired
    private AiService aiService;

    @Autowired
    private ExcelService excelService;

    @Autowired
    private AiExcelCommandParser aiExcelCommandParser;

    /**
     * 执行智能数据清理操作
     */
    public Map<String, Object> performSmartDataCleaning(MultipartFile file, String cleaningInstructions) throws IOException {
        logger.info("Starting smart data cleaning with instructions: {}", cleaningInstructions);
        
        Map<String, Object> result = new HashMap<>();
        
        try {
            // 验证输入
            if (file == null || file.isEmpty()) {
                logger.error("File is null or empty for data cleaning");
                result.put("success", false);
                result.put("error", "File is required for data cleaning");
                return result;
            }
            
            if (cleaningInstructions == null || cleaningInstructions.trim().isEmpty()) {
                logger.error("Cleaning instructions are null or empty");
                result.put("success", false);
                result.put("error", "Cleaning instructions are required");
                return result;
            }

            // 加载工作簿
            Workbook workbook = excelService.loadWorkbook(file);
            String excelData = excelService.getExcelDataAsString(workbook);

            // 构建AI请求
            AiRequest aiRequest = new AiRequest();
            aiRequest.setMessages(Arrays.asList(
                new AiRequest.Message("system",
                    "You are an Excel data cleaning expert. Based on the user's cleaning instructions, " +
                    "provide specific Excel operations to clean the data. " +
                    "Use the following command format when appropriate: " +
                    "[SET_CELL:A1:New Value] to set cell A1 to 'New Value', " +
                    "[INSERT_ROW:3:value1,value2,value3] to insert a row at position 3 with these values, " +
                    "[INSERT_COLUMN:2:value1,value2,value3] to insert a column at position 2 with these values, " +
                    "[DELETE_ROW:5] to delete row 5, " +
                    "[DELETE_COLUMN:1] to delete column 1, " +
                    "[APPLY_FORMULA:A1:B1+C1] to apply the formula 'B1+C1' in cell A1. " +
                    "For data cleaning, suggest operations like removing duplicates, fixing formatting, " +
                    "standardizing text, handling missing values, etc. " +
                    "Provide the commands in the format above embedded in your response."),
                new AiRequest.Message("user",
                    "Here is the Excel data:\n\n" + excelData + "\n\n" +
                    "Cleaning instructions: " + cleaningInstructions + "\n\n" +
                    "Please provide specific Excel operations to clean this data using the command format mentioned in the system message.")
            ));

            logger.debug("Sending data cleaning request to AI service");
            
            AiResponse aiResponse = aiService.generateResponse(aiRequest);
            String aiResponseContent = aiResponse.getChoices()[0].getMessage().getContent();

            // 解析并执行AI返回的Excel操作命令
            var commandResults = aiExcelCommandParser.parseAndExecuteCommands(workbook, aiResponseContent);

            result.put("aiResponse", aiResponseContent);
            result.put("cleaningInstructions", cleaningInstructions);
            result.put("commandResults", commandResults);
            result.put("success", true);
            
            logger.info("Smart data cleaning completed successfully");
            
        } catch (IOException e) {
            logger.error("IO error during smart data cleaning: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "IO error occurred during data cleaning: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Error during smart data cleaning: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "Error occurred during data cleaning: " + e.getMessage());
        }

        return result;
    }

    /**
     * 执行智能数据转换操作
     */
    public Map<String, Object> performSmartDataTransformation(MultipartFile file, String transformationInstructions) throws IOException {
        logger.info("Starting smart data transformation with instructions: {}", transformationInstructions);
        
        Map<String, Object> result = new HashMap<>();
        
        try {
            // 验证输入
            if (file == null || file.isEmpty()) {
                logger.error("File is null or empty for data transformation");
                result.put("success", false);
                result.put("error", "File is required for data transformation");
                return result;
            }
            
            if (transformationInstructions == null || transformationInstructions.trim().isEmpty()) {
                logger.error("Transformation instructions are null or empty");
                result.put("success", false);
                result.put("error", "Transformation instructions are required");
                return result;
            }

            // 加载工作簿
            Workbook workbook = excelService.loadWorkbook(file);
            String excelData = excelService.getExcelDataAsString(workbook);

            // 构建AI请求
            AiRequest aiRequest = new AiRequest();
            aiRequest.setMessages(Arrays.asList(
                new AiRequest.Message("system",
                    "You are an Excel data transformation expert. Based on the user's transformation instructions, " +
                    "provide specific Excel operations to transform the data. " +
                    "Use the following command format when appropriate: " +
                    "[SET_CELL:A1:New Value] to set cell A1 to 'New Value', " +
                    "[INSERT_ROW:3:value1,value2,value3] to insert a row at position 3 with these values, " +
                    "[INSERT_COLUMN:2:value1,value2,value3] to insert a column at position 2 with these values, " +
                    "[DELETE_ROW:5] to delete row 5, " +
                    "[DELETE_COLUMN:1] to delete column 1, " +
                    "[APPLY_FORMULA:A1:B1+C1] to apply the formula 'B1+C1' in cell A1. " +
                    "For data transformation, suggest operations like pivoting, merging, splitting, " +
                    "aggregating, calculating derived columns, etc. " +
                    "Provide the commands in the format above embedded in your response."),
                new AiRequest.Message("user",
                    "Here is the Excel data:\n\n" + excelData + "\n\n" +
                    "Transformation instructions: " + transformationInstructions + "\n\n" +
                    "Please provide specific Excel operations to transform this data using the command format mentioned in the system message.")
            ));

            logger.debug("Sending data transformation request to AI service");
            
            AiResponse aiResponse = aiService.generateResponse(aiRequest);
            String aiResponseContent = aiResponse.getChoices()[0].getMessage().getContent();

            // 解析并执行AI返回的Excel操作命令
            var commandResults = aiExcelCommandParser.parseAndExecuteCommands(workbook, aiResponseContent);

            result.put("aiResponse", aiResponseContent);
            result.put("transformationInstructions", transformationInstructions);
            result.put("commandResults", commandResults);
            result.put("success", true);
            
            logger.info("Smart data transformation completed successfully");
            
        } catch (IOException e) {
            logger.error("IO error during smart data transformation: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "IO error occurred during data transformation: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Error during smart data transformation: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "Error occurred during data transformation: " + e.getMessage());
        }

        return result;
    }

    /**
     * 执行智能数据分析操作
     */
    public Map<String, Object> performSmartDataAnalysis(MultipartFile file, String analysisInstructions) throws IOException {
        logger.info("Starting smart data analysis with instructions: {}", analysisInstructions);
        
        Map<String, Object> result = new HashMap<>();
        
        try {
            // 验证输入
            if (file == null || file.isEmpty()) {
                logger.error("File is null or empty for data analysis");
                result.put("success", false);
                result.put("error", "File is required for data analysis");
                return result;
            }
            
            if (analysisInstructions == null || analysisInstructions.trim().isEmpty()) {
                logger.error("Analysis instructions are null or empty");
                result.put("success", false);
                result.put("error", "Analysis instructions are required");
                return result;
            }

            // 加载工作簿
            Workbook workbook = excelService.loadWorkbook(file);
            String excelData = excelService.getExcelDataAsString(workbook);

            // 构建AI请求
            AiRequest aiRequest = new AiRequest();
            aiRequest.setMessages(Arrays.asList(
                new AiRequest.Message("system",
                    "You are an Excel data analysis expert. Based on the user's analysis instructions, " +
                    "provide specific Excel operations to analyze the data. " +
                    "Use the following command format when appropriate: " +
                    "[SET_CELL:A1:New Value] to set cell A1 to 'New Value', " +
                    "[INSERT_ROW:3:value1,value2,value3] to insert a row at position 3 with these values, " +
                    "[INSERT_COLUMN:2:value1,value2,value3] to insert a column at position 2 with these values, " +
                    "[DELETE_ROW:5] to delete row 5, " +
                    "[DELETE_COLUMN:1] to delete column 1, " +
                    "[APPLY_FORMULA:A1:B1+C1] to apply the formula 'B1+C1' in cell A1. " +
                    "For data analysis, suggest operations like summary statistics, trend analysis, " +
                    "pivot tables, conditional formatting, chart creation, etc. " +
                    "Provide the commands in the format above embedded in your response."),
                new AiRequest.Message("user",
                    "Here is the Excel data:\n\n" + excelData + "\n\n" +
                    "Analysis instructions: " + analysisInstructions + "\n\n" +
                    "Please provide specific Excel operations to analyze this data using the command format mentioned in the system message.")
            ));

            logger.debug("Sending data analysis request to AI service");
            
            AiResponse aiResponse = aiService.generateResponse(aiRequest);
            String aiResponseContent = aiResponse.getChoices()[0].getMessage().getContent();

            // 解析并执行AI返回的Excel操作命令
            var commandResults = aiExcelCommandParser.parseAndExecuteCommands(workbook, aiResponseContent);

            result.put("aiResponse", aiResponseContent);
            result.put("analysisInstructions", analysisInstructions);
            result.put("commandResults", commandResults);
            result.put("success", true);
            
            logger.info("Smart data analysis completed successfully");
            
        } catch (IOException e) {
            logger.error("IO error during smart data analysis: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "IO error occurred during data analysis: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Error during smart data analysis: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "Error occurred during data analysis: " + e.getMessage());
        }

        return result;
    }

    /**
     * 执行智能图表创建操作
     */
    public Map<String, Object> performSmartChartCreation(MultipartFile file, String chartInstructions) throws IOException {
        logger.info("Starting smart chart creation with instructions: {}", chartInstructions);
        
        Map<String, Object> result = new HashMap<>();
        
        try {
            // 验证输入
            if (file == null || file.isEmpty()) {
                logger.error("File is null or empty for chart creation");
                result.put("success", false);
                result.put("error", "File is required for chart creation");
                return result;
            }
            
            if (chartInstructions == null || chartInstructions.trim().isEmpty()) {
                logger.error("Chart instructions are null or empty");
                result.put("success", false);
                result.put("error", "Chart instructions are required");
                return result;
            }

            // 加载工作簿
            Workbook workbook = excelService.loadWorkbook(file);
            String excelData = excelService.getExcelDataAsString(workbook);

            // 构建AI请求
            AiRequest aiRequest = new AiRequest();
            aiRequest.setMessages(Arrays.asList(
                new AiRequest.Message("system",
                    "You are an Excel chart creation expert. Based on the user's chart instructions, " +
                    "provide specific Excel operations to create charts. " +
                    "Use the following command format when appropriate: " +
                    "[SET_CELL:A1:New Value] to set cell A1 to 'New Value', " +
                    "[INSERT_ROW:3:value1,value2,value3] to insert a row at position 3 with these values, " +
                    "[INSERT_COLUMN:2:value1,value2,value3] to insert a column at position 2 with these values, " +
                    "[DELETE_ROW:5] to delete row 5, " +
                    "[DELETE_COLUMN:1] to delete column 1, " +
                    "[APPLY_FORMULA:A1:B1+C1] to apply the formula 'B1+C1' in cell A1. " +
                    "For chart creation, suggest operations to select data, insert charts, " +
                    "apply formatting, set axis labels, etc. " +
                    "Provide the commands in the format above embedded in your response."),
                new AiRequest.Message("user",
                    "Here is the Excel data:\n\n" + excelData + "\n\n" +
                    "Chart creation instructions: " + chartInstructions + "\n\n" +
                    "Please provide specific Excel operations to create the chart using the command format mentioned in the system message.")
            ));

            logger.debug("Sending chart creation request to AI service");
            
            AiResponse aiResponse = aiService.generateResponse(aiRequest);
            String aiResponseContent = aiResponse.getChoices()[0].getMessage().getContent();

            // 解析并执行AI返回的Excel操作命令
            var commandResults = aiExcelCommandParser.parseAndExecuteCommands(workbook, aiResponseContent);

            result.put("aiResponse", aiResponseContent);
            result.put("chartInstructions", chartInstructions);
            result.put("commandResults", commandResults);
            result.put("success", true);
            
            logger.info("Smart chart creation completed successfully");
            
        } catch (IOException e) {
            logger.error("IO error during smart chart creation: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "IO error occurred during chart creation: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Error during smart chart creation: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "Error occurred during chart creation: " + e.getMessage());
        }

        return result;
    }

    /**
     * 执行智能数据验证操作
     */
    public Map<String, Object> performSmartDataValidation(MultipartFile file, String validationInstructions) throws IOException {
        logger.info("Starting smart data validation with instructions: {}", validationInstructions);
        
        Map<String, Object> result = new HashMap<>();
        
        try {
            // 验证输入
            if (file == null || file.isEmpty()) {
                logger.error("File is null or empty for data validation");
                result.put("success", false);
                result.put("error", "File is required for data validation");
                return result;
            }
            
            if (validationInstructions == null || validationInstructions.trim().isEmpty()) {
                logger.error("Validation instructions are null or empty");
                result.put("success", false);
                result.put("error", "Validation instructions are required");
                return result;
            }

            // 加载工作簿
            Workbook workbook = excelService.loadWorkbook(file);
            String excelData = excelService.getExcelDataAsString(workbook);

            // 构建AI请求
            AiRequest aiRequest = new AiRequest();
            aiRequest.setMessages(Arrays.asList(
                new AiRequest.Message("system",
                    "You are an Excel data validation expert. Based on the user's validation instructions, " +
                    "provide specific Excel operations to validate the data. " +
                    "Use the following command format when appropriate: " +
                    "[SET_CELL:A1:New Value] to set cell A1 to 'New Value', " +
                    "[INSERT_ROW:3:value1,value2,value3] to insert a row at position 3 with these values, " +
                    "[INSERT_COLUMN:2:value1,value2,value3] to insert a column at position 2 with these values, " +
                    "[DELETE_ROW:5] to delete row 5, " +
                    "[DELETE_COLUMN:1] to delete column 1, " +
                    "[APPLY_FORMULA:A1:B1+C1] to apply the formula 'B1+C1' in cell A1. " +
                    "For data validation, suggest operations like setting data validation rules, " +
                    "applying conditional formatting for errors, identifying outliers, etc. " +
                    "Provide the commands in the format above embedded in your response."),
                new AiRequest.Message("user",
                    "Here is the Excel data:\n\n" + excelData + "\n\n" +
                    "Validation instructions: " + validationInstructions + "\n\n" +
                    "Please provide specific Excel operations to validate this data using the command format mentioned in the system message.")
            ));

            logger.debug("Sending data validation request to AI service");
            
            AiResponse aiResponse = aiService.generateResponse(aiRequest);
            String aiResponseContent = aiResponse.getChoices()[0].getMessage().getContent();

            // 解析并执行AI返回的Excel操作命令
            var commandResults = aiExcelCommandParser.parseAndExecuteCommands(workbook, aiResponseContent);

            result.put("aiResponse", aiResponseContent);
            result.put("validationInstructions", validationInstructions);
            result.put("commandResults", commandResults);
            result.put("success", true);
            
            logger.info("Smart data validation completed successfully");
            
        } catch (IOException e) {
            logger.error("IO error during smart data validation: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "IO error occurred during data validation: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Error during smart data validation: {}", e.getMessage(), e);
            result.put("success", false);
            result.put("error", "Error occurred during data validation: " + e.getMessage());
        }

        return result;
    }
}