package com.example.aiexcel.unit.service;

import com.example.aiexcel.service.AiExcelCommandParser;
import com.example.aiexcel.service.excel.ExcelService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class AiExcelCommandParserTest {

    @Mock
    private ExcelService excelService;

    @InjectMocks
    private AiExcelCommandParser commandParser;

    private Workbook testWorkbook;

    @BeforeEach
    void setUp() {
        testWorkbook = new XSSFWorkbook();
        Sheet sheet = testWorkbook.createSheet("Sheet1");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Value");
        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue("Test");
        dataRow.createCell(1).setCellValue(100);
    }

    @Test
    void testParseAndExecuteCommands_EmptyResponse() {
        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, "");

        assertTrue(results.isEmpty());
    }

    @Test
    void testParseAndExecuteCommands_NullResponse() {
        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, null);

        assertTrue(results.isEmpty());
    }

    @Test
    void testParseAndExecuteCommands_SetCell() {
        String aiResponse = "Setting cell value [SET_CELL:A1:NewValue]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertFalse(results.isEmpty());
        assertTrue(results.get(0).isSuccess());
        assertEquals("SET_CELL", results.get(0).getCommandType());
        verify(excelService, times(1)).updateCell(eq(testWorkbook), anyString(), eq(0), eq(0), eq("NewValue"));
    }

    @Test
    void testParseAndExecuteCommands_InsertRow() {
        String aiResponse = "Inserting row [INSERT_ROW:2:val1,val2,val3]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertFalse(results.isEmpty());
        assertEquals("INSERT_ROW", results.get(0).getCommandType());
    }

    @Test
    void testParseAndExecuteCommands_ApplyFormula() {
        String aiResponse = "Applying formula [APPLY_FORMULA:C1:5+10]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertFalse(results.isEmpty());
        assertEquals("APPLY_FORMULA", results.get(0).getCommandType());
    }

    @Test
    void testParseAndExecuteCommands_DeleteRow() {
        String aiResponse = "Deleting row [DELETE_ROW:1]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertFalse(results.isEmpty());
        assertEquals("DELETE_ROW", results.get(0).getCommandType());
    }

    @Test
    void testParseAndExecuteCommands_DeleteColumn() {
        String aiResponse = "Deleting column [DELETE_COLUMN:0]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertFalse(results.isEmpty());
        assertEquals("DELETE_COLUMN", results.get(0).getCommandType());
    }

    @Test
    void testParseAndExecuteCommands_InsertColumn() {
        String aiResponse = "Inserting column [INSERT_COLUMN:1:cval1,cval2]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertFalse(results.isEmpty());
        assertEquals("INSERT_COLUMN", results.get(0).getCommandType());
    }

    @Test
    void testParseAndExecuteCommands_MultipleCommands() {
        String aiResponse = "[SET_CELL:A1:Hello] and [SET_CELL:B1:World]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertEquals(2, results.size());
        assertTrue(results.get(0).isSuccess());
        assertTrue(results.get(1).isSuccess());
        verify(excelService, times(2)).updateCell(eq(testWorkbook), anyString(), anyInt(), anyInt(), anyString());
    }

    @Test
    void testParseAndExecuteCommands_InvalidCellRef() {
        String aiResponse = "[SET_CELL:123:InvalidRef]";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertTrue(results.isEmpty(), "Invalid cell reference should be skipped");
    }

    @Test
    void testParseAndExecuteCommands_NoCommands() {
        String aiResponse = "This is just a normal AI response without any commands.";

        List<AiExcelCommandParser.CommandResult> results =
                commandParser.parseAndExecuteCommands(testWorkbook, aiResponse);

        assertTrue(results.isEmpty());
    }
}
