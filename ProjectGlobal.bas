Attribute VB_Name = "ProjectGlobal"
Option Explicit

Public Enum ProjectInfoColumns
    [_First] = 1                                 'columns numbers are relative to the anchor cell
    ProjectNumber = 1
    SelectCheckBox = 2
    DBName = 3
    FullPath = 4
    BrowseButton = 5
    FileTimestamp = 6
    SelectedCellLink = 7
    [_Last] = 7
End Enum

Public Enum ProjectApp_Errors
    NoProjectArea = vbObject + 600
    IndexOutOfBounds
End Enum

Public Const DEFAULT_PROJECT_ROWS As Long = 15
Public Const ROW_ONE_ANCHOR As String = "A3"

Public Function ConnectToModel() As ProjectsModel
    Dim anchor As Range
    Dim projectArea As Range
    Set anchor = ThisWorkbook.Sheets("Configure").Range(ROW_ONE_ANCHOR)
    Set projectArea = anchor.Resize(DEFAULT_PROJECT_ROWS, ProjectInfoColumns.[_Last])

    Dim pvm As ProjectsModel
    Set pvm = New ProjectsModel
    pvm.Connect projectArea
    Set ConnectToModel = pvm
End Function

Public Function ProjectEditableArea() As Range
    Dim anchor As Range
    Dim projectArea As Range
    Set anchor = ThisWorkbook.Sheets("Configure").Range(ROW_ONE_ANCHOR)
    Set projectArea = anchor.Resize(DEFAULT_PROJECT_ROWS, ProjectInfoColumns.[_Last])

    Set ProjectEditableArea = Application.Union(projectArea.Columns(ProjectInfoColumns.DBName), _
                                                projectArea.Columns(ProjectInfoColumns.FullPath))
End Function

Public Function ProjectSelectAllCell() As Range
    Dim anchor As Range
    Set anchor = ThisWorkbook.Sheets("Configure").Range(ROW_ONE_ANCHOR)
    Set ProjectSelectAllCell = anchor.Offset(-1, ProjectInfoColumns.SelectedCellLink - 1)
End Function

Public Function ProjectInfoColumnsName(ByVal index As Long) As String
    ProjectInfoColumnsName = vbNullString
    If (index >= ProjectInfoColumns.[_First]) Or _
       (index <= ProjectInfoColumns.[_Last]) Then
       
        Dim names() As String
        names = Split(",Name in DB,MS Excel Path,," & _
                      "MS Excel File Timestamp,,", _
                      ",", , vbTextCompare)
    
        ProjectInfoColumnsName = names(index - 1)
    End If
End Function

Public Function ProjectRowIndex(ByVal wsRowIndex As Long) As Long
    '--- given the worksheet row number, this returns the project
    '    index row into the group of projects
    Dim anchor As Range
    Set anchor = ThisWorkbook.Sheets("Configure").Range(ROW_ONE_ANCHOR)
    ProjectRowIndex = wsRowIndex - anchor.Row + 1
End Function


