package com.example.aiexcel.service.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

public interface ExcelService {
    Workbook loadWorkbook(MultipartFile file) throws IOException;
    Workbook loadWorkbook(InputStream inputStream) throws IOException;
    void saveWorkbook(Workbook workbook, String filePath) throws IOException;
    byte[] getWorkbookAsBytes(Workbook workbook) throws IOException;
    String getExcelDataAsString(Workbook workbook);
    Object[][] getExcelDataAsArray(Workbook workbook);
    void updateCell(Workbook workbook, String sheetName, int rowIndex, int colIndex, Object value);
    void updateRange(Workbook workbook, String sheetName, int startRowIndex, int startColIndex,
                     int endRowIndex, int endColIndex, Object[][] values);

    // Enhanced Excel processing capabilities
    void applyCellFormatting(Workbook workbook, String sheetName, int rowIndex, int colIndex,
                             String format, String color);
    void insertRow(Workbook workbook, String sheetName, int rowIndex, Object[] values);
    void insertColumn(Workbook workbook, String sheetName, int colIndex, Object[] values);
    void deleteRow(Workbook workbook, String sheetName, int rowIndex);
    void deleteColumn(Workbook workbook, String sheetName, int colIndex);
    void applyFormula(Workbook workbook, String sheetName, int rowIndex, int colIndex, String formula);
    String getCellValue(Workbook workbook, String sheetName, int rowIndex, int colIndex);
    int getRowCount(Workbook workbook, String sheetName);
    int getColumnCount(Workbook workbook, String sheetName, int rowIndex);
    String[] getExcelHeaders(Workbook workbook);
    void evaluateAllFormulasInWorkbook(Workbook workbook);
}