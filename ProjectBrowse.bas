Attribute VB_Name = "ProjectBrowse"
Option Explicit

Public Sub UserFileSelect()
    '--- connected to the "..." button on the Configure worksheet to select
    '    an MS Project file for that row
    Dim wsRow As Long
    Dim projectIndex As Long
    wsRow = ThisWorkbook.Sheets("Configure").Shapes(Application.Caller).TopLeftCell.Row
    projectIndex = ProjectRowIndex(wsRow)
    
    Dim pvm As ProjectsModel
    Set pvm = ConnectToModel()
    
    Dim projInfo As ProjectInfo
    Set projInfo = pvm.GetProject(projectIndex)

    Dim filePicker As Office.FileDialog
    Set filePicker = Application.FileDialog(MsoFileDialogType.msoFileDialogOpen)
    With filePicker
        .Title = "Select an MS Excel File to Add..."
        .Filters.Add "MS Project", "*.xlsx", 1
        .AllowMultiSelect = False
        If Len(projInfo.FullPath) > 0 Then
            .InitialFileName = projInfo.FullPath
        Else
            .InitialFileName = ThisWorkbook.path
        End If
        If .Show Then
            projInfo.FullPath = .SelectedItems(1)
        End If
    End With
End Sub

