Attribute VB_Name = "CreateProjectsSheet"
Option Explicit

Public Sub BuildProjectView()
    '--- wipe the sheet and build it up to ensure everything is
    '    in the right place
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Configure")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Configure"
    End If
    ResetSheet ws
    
    Dim anchor As Range
    Dim projectArea As Range
    Dim headerArea As Range
    Dim importArea As Range
    Set anchor = ws.Range(ROW_ONE_ANCHOR)
    Set projectArea = anchor.Resize(DEFAULT_PROJECT_ROWS, ProjectInfoColumns.[_Last])
    Set headerArea = anchor.Offset(-1, 1).Resize(1, ProjectInfoColumns.[_Last] - 2)
    Set importArea = anchor.Offset(DEFAULT_PROJECT_ROWS + 1, 2)
    
    ws.Cells.Interior.color = XlRgbColor.rgbPaleTurquoise  'overall bg color
    
    FormatProjectArea projectArea, headerArea
    CreateBrowseButtons ws, projectArea
    CreateCheckboxes ws, projectArea
    CreateImportButton ws, importArea
    
    anchor.Offset(, 2).Select
End Sub

Private Sub ResetSheet(ByRef ws As Worksheet)
    With ws
        .Cells.Clear
        .Cells.ColumnWidth = 8.14
        .Cells.EntireRow.AutoFit
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        
        Do While .CheckBoxes.Count > 0
            .CheckBoxes(1).Delete
        Loop
        
        Do While .Buttons.Count > 0
            .Buttons(1).Delete
        Loop
        
        Do While .OLEObjects.Count > 0
            .OLEObjects(1).Delete
        Loop
    End With
End Sub

Private Sub FormatProjectArea(ByRef projectArea As Range, _
                              ByRef headerArea As Range)
    Dim areaWithoutNumbers As Range
    Set areaWithoutNumbers = projectArea.Offset(0, 1).Resize(, ProjectInfoColumns.[_Last] - 2)
    With areaWithoutNumbers.Borders
        .LineStyle = xlContinuous
        .color = XlRgbColor.rgbDarkGray
        .Weight = xlThin
    End With
    
    '--- predefined column widths and formats
    With projectArea
        .Columns(ProjectInfoColumns.ProjectNumber).ColumnWidth = 3#
        .Columns(ProjectInfoColumns.SelectCheckBox).ColumnWidth = 3#
        .Columns(ProjectInfoColumns.DBName).ColumnWidth = 11#
        .Columns(ProjectInfoColumns.FullPath).ColumnWidth = 40#
        .Columns(ProjectInfoColumns.BrowseButton).ColumnWidth = 5#
        .Columns(ProjectInfoColumns.FileTimestamp).ColumnWidth = 14#
        
        .Cells(1, ProjectInfoColumns.SelectCheckBox).Resize(.Rows.Count, 1).Interior.color = rgbWhite
        .Cells(1, ProjectInfoColumns.DBName).Resize(.Rows.Count, 1).Interior.color = rgbWhite
        .Cells(1, ProjectInfoColumns.FullPath).Resize(.Rows.Count, 1).Interior.color = rgbWhite
        
        .Cells(1, ProjectInfoColumns.FileTimestamp).Resize(.Rows.Count, 1).NumberFormat = "dd-mmm-yyyy"
        
        '--- the linked cell needs to have the same font color as the background
        .Cells(1, ProjectInfoColumns.SelectedCellLink).Resize(.Rows.Count, 1).Font.color = XlRgbColor.rgbPaleTurquoise
    End With
    
    With headerArea
        .Cells.Interior.color = XlRgbColor.rgbDarkBlue
        .Cells.Font.color = XlRgbColor.rgbWhite
        .Cells.Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.color = XlRgbColor.rgbDarkGray
        .Borders.Weight = xlThin
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .EntireRow.AutoFit
        
        Dim i As Long
        '--- start at _First+1 to skip the ProjectNumber column
        For i = (ProjectInfoColumns.[_First] + 1) To ProjectInfoColumns.[_Last]
            .Cells(1, i).Value = ProjectInfoColumnsName(i)
        Next i
    
        '--- the linked cell needs to have the same font color as the background
        .Cells(1, .Columns.Count + 1).Font.color = XlRgbColor.rgbPaleTurquoise
    End With
    
    '--- label the project indexes
    With projectArea
        For i = 1 To .Rows.Count
            .Cells(i, ProjectInfoColumns.ProjectNumber).Value = i
        Next i
        .Resize(.Rows.Count, 1).Font.Size = 8
    End With
End Sub

Private Sub CreateBrowseButtons(ByRef ws As Worksheet, ByRef projectArea As Range)
    Dim projectRow As Range
    Set projectRow = projectArea.Resize(1, ProjectInfoColumns.[_Last])
    
    Dim left As Double
    Dim top As Double
    Dim height As Double
    Dim width As Double
    
    Dim i As Long
    For i = 1 To DEFAULT_PROJECT_ROWS
        With projectRow
            height = .height - 2
            width = .Cells(1, ProjectInfoColumns.BrowseButton).width - 2
            top = .top + 1
            left = .Cells(1, ProjectInfoColumns.BrowseButton).left + 1
            
            Dim btn As Button
            Set btn = ws.Buttons.Add(left:=left, top:=top, _
                                     height:=height, width:=width)
            btn.Caption = "..."
            btn.Enabled = True
            btn.OnAction = "UserFileSelect"
            
            Set projectRow = projectRow.Offset(1, 0)
        End With
    Next i
End Sub

Private Sub CreateCheckboxes(ByRef ws As Worksheet, ByRef projectArea As Range)
    Dim projectRow As Range
    Set projectRow = projectArea.Resize(1, ProjectInfoColumns.[_Last] + 1)
    
    Dim left As Double
    Dim top As Double
    Dim height As Double
    Dim width As Double
    
    Dim cb As CheckBox
    Dim i As Long
    For i = 1 To DEFAULT_PROJECT_ROWS
        With projectRow
            top = .top + 1
            height = .height - 2
            width = 14
            With .Cells(1, ProjectInfoColumns.SelectCheckBox)
                '--- centered in the column
                left = .left + (.width / 2#) - (width / 2#)
            End With
        End With
        Set cb = ws.CheckBoxes.Add(left:=left, top:=top, _
                                   height:=height, width:=width)
        cb.Caption = vbNullString
        cb.LinkedCell = projectRow.Cells(1, ProjectInfoColumns.SelectedCellLink).Address
        projectRow.Cells(1, ProjectInfoColumns.SelectedCellLink).Value = False
        
        '--- sometimes the height and width don't get correctly set during
        '    the Add call, so set them (again) here, with other settings
        cb.height = height
        cb.width = width
            
        Set projectRow = projectRow.Offset(1, 0)
    Next i
        
    '--- now add the Select All checkbox above these
    '    keep the same height, width, and left position, but
    '    recalculate the Top to set it at the botton of the cell
    Set projectRow = projectArea.Offset(-1, 0).Resize(1, ProjectInfoColumns.[_Last] + 1)
    top = projectRow.top + projectRow.height - height - 2
    Set cb = ws.CheckBoxes.Add(left:=left, top:=top, _
                               height:=height, width:=width)
    cb.Caption = vbNullString
    cb.OnAction = "CheckUncheckAll"
    cb.LinkedCell = projectRow.Cells(1, ProjectInfoColumns.SelectedCellLink).Address
    projectRow.Cells(1, ProjectInfoColumns.SelectedCellLink).Value = False
End Sub

Private Sub CreateImportButton(ByRef ws As Worksheet, _
                               ByRef importArea As Range)
    Dim left As Double
    Dim top As Double
    Dim height As Double
    Dim width As Double
    
    With importArea
        height = .height * 2#
        width = .width * 2.5
        left = .left
        top = .top
        
        Dim btn As Button
        Set btn = ws.Buttons.Add(left:=left, top:=top, _
                                 height:=height, width:=width)
        btn.Caption = "Import Select Projects"
        btn.Enabled = True
        btn.OnAction = "ImportProjects"
    End With
End Sub

