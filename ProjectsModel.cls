VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type InternalData
    storage As Range
    projects As Collection    'collection of ProjectInfo objects
End Type
Private this As InternalData

Public Property Get IsConnected() As Boolean
    IsConnected = (Not this.projects Is Nothing)
End Property

Public Property Get ProjectCount() As Long
    If Not this.projects Is Nothing Then
        ProjectCount = this.projects.Count
    Else
        ProjectCount = 0
    End If
End Property

Public Property Get GetProject(ByVal index As Long) As ProjectInfo
    If Not this.projects Is Nothing Then
        If index > 0 And index <= this.projects.Count Then
            Set GetProject = this.projects(index)
        Else
            Err.Raise ProjectApp_Errors.IndexOutOfBounds, _
                      "ProjectsModel::GetProject", _
                      "Project index out of bounds"
        End If
    Else
        Set GetProject = Nothing
    End If
End Property

Public Sub Connect(ByRef projectArea As Range)
    If this.storage Is Nothing Then
        If Not projectArea Is Nothing Then
            Set this.storage = projectArea
        Else
            Err.Raise ProjectApp_Errors.NoProjectArea, _
                      "ProjectsModel::Load", _
                      "Missing project area"
        End If
    ElseIf Not this.storage = projectArea Then
        '--- we've got a new table, so dump the old one and reload
        Set this.storage = Nothing
        Set this.storage = projectArea
    End If
    
    Set this.projects = Nothing
    Set this.projects = New Collection
    
    Dim projectRow As Range
    Set projectRow = this.storage.Resize(1, ProjectInfoColumns.[_Last])

    Dim i As Long
    Dim newInfo As ProjectInfo
    For i = 1 To projectArea.Rows.Count
        Set newInfo = New ProjectInfo
        newInfo.Connect projectRow
        this.projects.Add newInfo
        Set projectRow = projectRow.Offset(1, 0)
    Next i
End Sub

Public Function GetSelectedProjects() As Collection
    Dim selectedCollection As Collection
    Dim projectRow As Variant
    For Each projectRow In this.projects
        If projectRow.IsSelected Then
            If selectedCollection Is Nothing Then
                Set selectedCollection = New Collection
            End If
            selectedCollection.Add projectRow
        End If
    Next projectRow
    Set GetSelectedProjects = selectedCollection
End Function

Private Sub Class_Initialize()

End Sub

