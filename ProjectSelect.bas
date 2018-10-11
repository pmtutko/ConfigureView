Attribute VB_Name = "ProjectSelect"
Option Explicit

Public Sub CheckUncheckAll()
    Dim selectAllBox As Range
    Set selectAllBox = ProjectSelectAllCell()
    
    Dim pvm As ProjectsModel
    Dim projInfo As ProjectInfo
    Set pvm = ConnectToModel()
        
    Dim i As Long
    For i = 1 To pvm.ProjectCount
        Set projInfo = pvm.GetProject(i)
        projInfo.IsSelected = selectAllBox
    Next i
End Sub
