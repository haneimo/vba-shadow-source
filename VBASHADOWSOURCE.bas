Attribute VB_Name = "VBASHADOWSOURCE"
Option Explicit

'--
' エクスポートソース
' 参考：http://www.excellenceweb.net/vba/programing/export.html
'--
Private Enum ComponentType
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
End Enum

Public Sub exportVBAFiles()
    Dim tempComponent       As Object
    Dim ExportPath          As String
    ExportPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".", "_") & "_src"
    
    If Dir(ExportPath, vbDirectory) = "" Then
        Call MkDir(ExportPath)
    End If

    For Each tempComponent In ThisWorkbook.VBProject.VBComponents
        If tempComponent.Name = "VBASHADOWSOURCE" Then GoTo NEXTLOOP
        Select Case tempComponent.Type
            Case STANDARD_MODULE
                tempComponent.Export ExportPath & "\" & tempComponent.Name & ".bas"
            Case CLASS_MODULE
                tempComponent.Export ExportPath & "\" & tempComponent.Name & ".cls"
            Case USER_FORM
                tempComponent.Export ExportPath & "\" & tempComponent.Name & ".frm"
        End Select
NEXTLOOP:
    Next
End Sub

'--
' インポートソース
'--
Public Sub importVBAFiles()
    Dim sourceFile As String
    Dim ExportPath As String
    Dim ext As String
    ExportPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".", "_") & "_src"
    
    If Dir(ExportPath, vbDirectory) = "" Then
        Call MkDir(ExportPath)
    End If
    
    Dim tempComponent As Object
    For Each tempComponent In ThisWorkbook.VBProject.VBComponents
        If tempComponent.Name = "VBASHADOWSOURCE" Then GoTo NEXTLOOP
        
        If tempComponent.Type = STANDARD_MODULE Or tempComponent.Type = CLASS_MODULE Or tempComponent.Type = USER_FORM Then
            Call ThisWorkbook.VBProject.VBComponents.Remove(tempComponent)
        End If
NEXTLOOP:
    Next
    
    sourceFile = Dir(ExportPath & "\*", vbNormal)
    While sourceFile <> ""
        ext = Right(sourceFile, 4)
        If ext = ".bas" Or ext = ".cls" Or ext = ".frm" Then
            ThisWorkbook.VBProject.VBComponents.Import ExportPath & "\" & sourceFile
        End If
        sourceFile = Dir()
    Wend
End Sub


Private Sub initVBASHADOWSOURCE()
    Dim code As String
    Const INSERTEDSYMBOL = "'VBASHADOWSOURE-INSERTED"
    code = INSERTEDSYMBOL & vbCrLf _
         & "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)" & vbCrLf _
         & "  'コードをエクスポート" & vbCrLf _
         & "  Call exportVBAFiles" & vbCrLf _
         & "End Sub" & vbCrLf _
         & "Private Sub Workbook_Open()" & vbCrLf _
         & " 'コードをインポート" & vbCrLf _
         & "  Call importVBAFiles" & vbCrLf _
         & " End Sub"
    
    Dim tempComponent As Object
    For Each tempComponent In ThisWorkbook.VBProject.VBComponents
        If tempComponent.Name = "ThisWorkbook" Then
            If Not tempComponent.CodeModule.Find(INSERTEDSYMBOL, 0, 0, 0, 0) Then
                tempComponent.CodeModule.InsertLines 1, code
            End If
        End If
    Next
End Sub

