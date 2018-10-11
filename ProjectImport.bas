Attribute VB_Name = "ProjectImport"
Option Explicit

Public Sub ImportProjects()
    Dim pvm As ProjectsModel
    Set pvm = ConnectToModel()
    
    Dim selectedProjects As Collection
    Set selectedProjects = pvm.GetSelectedProjects
    
    Dim text As String
    If selectedProjects Is Nothing Then
        text = "No projects selected!"
    Else
        Dim proj As Variant
        For Each proj In selectedProjects
            text = text & " -- " & proj.Filename & vbCrLf
        Next proj
    End If
    
    MsgBox "You selected these projects for import:" & vbCrLf & text, _
           vbInformation + vbOKOnly
End Sub

