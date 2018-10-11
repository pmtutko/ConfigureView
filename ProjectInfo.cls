VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type InternalData
    projectRow As Range
    selectBox As CheckBox
End Type
Private this As InternalData

Public Property Let IsSelected(ByVal newState As Boolean)
    '--- can only be selected if there is a valid filename
    If Len(this.projectRow.Cells(1, ProjectInfoColumns.FullPath)) > 0 Then
        this.projectRow.Cells(1, ProjectInfoColumns.SelectedCellLink).Value = newState
    End If
End Property

Public Property Get IsSelected() As Boolean
    IsSelected = this.projectRow.Cells(1, ProjectInfoColumns.SelectedCellLink).Value
End Property

Public Property Let NameInDB(ByVal newName As String)
    '--- disable events because this area is watched by a
    '    Worksheet_Change event
    Application.EnableEvents = False
    this.projectRow.Cells(1, ProjectInfoColumns.DBName).Value = newName
    Application.EnableEvents = True
End Property

Public Property Get NameInDB() As String
    NameInDB = this.projectRow.Cells(1, ProjectInfoColumns.DBName).Value
End Property

Public Property Let FullPath(ByVal newPath As String)
    '--- disable events because this area is watched by a
    '    Worksheet_Change event
    Application.EnableEvents = False
    If Len(newPath) > 0 Then
        this.projectRow.Cells(1, ProjectInfoColumns.FullPath).Value = newPath
        On Error GoTo TimestampError
        Me.FileTimestamp = FileDateTime(newPath)
    Else
        '--- clear all the data if the user deleted the name
        this.projectRow.Cells(1, ProjectInfoColumns.DBName).Value = vbNullString
        this.projectRow.Cells(1, ProjectInfoColumns.FullPath).Value = vbNullString
        this.projectRow.Cells(1, ProjectInfoColumns.FileTimestamp).Value = vbNullString
    End If
    
NormalExit:
    SetCheckBoxState
    Application.EnableEvents = True
    Exit Property
TimestampError:
    If Err.Number = 53 Then
        MsgBox "The filename you entered is not valid. Please enter " & _
               "a valid filename.", vbCritical + vbOKOnly, _
               "Error In Filename"
        this.projectRow.Cells(1, ProjectInfoColumns.FullPath).Value = vbNullString
    End If
    GoTo NormalExit
End Property

Public Property Get FullPath() As String
    FullPath = this.projectRow.Cells(1, ProjectInfoColumns.FullPath)
End Property

Public Property Get PathOnly() As String
    '--- returns ONLY the path part of the full path
    PathOnly = vbNullString
    If Len(Me.FullPath) = 0 Then
        Exit Property
    End If
    
    Dim pos1 As Long
    pos1 = InStrRev(Me.FullPath, "\", , vbTextCompare)
    If pos1 > 0 Then
        PathOnly = left$(Me.FullPath, pos1 - 1)
    Else
        PathOnly = Me.FullPath
    End If
End Property

Public Property Get Filename() As String
    '--- returns ONLY the filename part of the full path
    Filename = vbNullString
    If Len(Me.FullPath) = 0 Then
        Exit Property
    End If
    
    Dim pos1 As Long
    pos1 = InStrRev(Me.FullPath, "\", , vbTextCompare)
    If pos1 > 0 Then
        Filename = Right$(Me.FullPath, Len(Me.FullPath) - pos1)
    Else
        Filename = Me.FullPath
    End If
End Property

Public Property Let FileTimestamp(ByVal newDate As Date)
    this.projectRow.Cells(1, ProjectInfoColumns.FileTimestamp).Value = newDate
End Property

Public Property Get FileTimestamp() As Date
    FileTimestamp = this.projectRow.Cells(1, ProjectInfoColumns.FileTimestamp).Value
End Property

Public Sub Connect(ByRef dataRow As Range)
    '--- establishes the connection of this object to the given worksheet
    '    row of project data
    Set this.projectRow = dataRow
    FindMyCheckBox
    SetCheckBoxState
End Sub

Private Sub Class_Initialize()

End Sub

Private Sub FindMyCheckBox()
    '--- quickly loop through all the checkboxes on the worksheet and find the
    '    one that is located in this assigned row. there can only be one
    Dim cb As CheckBox
    Dim ws As Worksheet
    Set ws = this.projectRow.Parent
    For Each cb In ws.CheckBoxes
        If (cb.top >= this.projectRow.top) And _
           (cb.top < (this.projectRow.top + this.projectRow.height)) Then
            Set this.selectBox = cb
            Exit For
        End If
    Next cb
End Sub

Private Sub SetCheckBoxState()
    '--- the checkbox is enabled if there is a valid filename
    If Not this.selectBox Is Nothing Then
        If Len(Me.FullPath) > 0 Then
            this.selectBox.Enabled = True
        Else
            this.selectBox.Value = False
            this.selectBox.Enabled = False
        End If
    End If
End Sub


