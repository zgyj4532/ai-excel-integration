package com.example.aiexcel.unit.service;

import com.example.aiexcel.model.FileWorkspace;
import com.example.aiexcel.repository.WorkspaceFileRepository;
import com.example.aiexcel.repository.WorkspaceRepository;
import com.example.aiexcel.service.FileManagerService;
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
class FileManagerServiceTest {

    @Mock
    private WorkspaceRepository workspaceRepository;

    @Mock
    private WorkspaceFileRepository workspaceFileRepository;

    @InjectMocks
    private FileManagerService fileManagerService;

    @Test
    void testCreateWorkspace_Success() {
        String name = "My Workspace";
        String userId = "user1";
        when(workspaceRepository.existsByUserIdAndName(userId, name)).thenReturn(false);
        when(workspaceRepository.save(any(FileWorkspace.class)))
                .thenAnswer(invocation -> {
                    FileWorkspace ws = invocation.getArgument(0);
                    ws.setId(1L);
                    return ws;
                });

        FileWorkspace result = fileManagerService.createWorkspace(name, "desc", userId, null);

        assertNotNull(result);
        assertEquals(name, result.getName());
        assertEquals(userId, result.getUserId());
        verify(workspaceRepository, times(1)).save(any(FileWorkspace.class));
    }

    @Test
    void testCreateWorkspace_DuplicateName() {
        String name = "Existing";
        String userId = "user1";
        when(workspaceRepository.existsByUserIdAndName(userId, name)).thenReturn(true);

        assertThrows(IllegalArgumentException.class,
                () -> fileManagerService.createWorkspace(name, "desc", userId, null));

        verify(workspaceRepository, never()).save(any());
    }

    @Test
    void testGetUserWorkspaces_ReturnsList() {
        String userId = "user1";
        FileWorkspace ws = new FileWorkspace("WS1", "desc", userId, null);
        when(workspaceRepository.findByUserId(userId)).thenReturn(List.of(ws));

        List<FileWorkspace> result = fileManagerService.getUserWorkspaces(userId);

        assertEquals(1, result.size());
        assertEquals("WS1", result.get(0).getName());
    }

    @Test
    void testGetUserWorkspaces_Empty() {
        when(workspaceRepository.findByUserId("no-user")).thenReturn(List.of());

        List<FileWorkspace> result = fileManagerService.getUserWorkspaces("no-user");

        assertTrue(result.isEmpty());
    }

    @Test
    void testGetChildWorkspaces() {
        Long parentId = 1L;
        FileWorkspace child = new FileWorkspace("Child", "desc", "user1", parentId);
        when(workspaceRepository.findByParentWorkspaceId(parentId)).thenReturn(List.of(child));

        List<FileWorkspace> result = fileManagerService.getChildWorkspaces(parentId);

        assertEquals(1, result.size());
    }
}
