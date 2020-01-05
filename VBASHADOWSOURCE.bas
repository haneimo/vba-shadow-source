Attribute VB_Name = "VBASHADOWSOURCE"
Option Explicit

Private Const OWNMODULENAME As String = "VBASHADOWSOURCE"
Private Const InsertedStartSymbol As String = "'VBASHADOWSOURE-INSERTED-START"
Private Const InsertedEndSymbol As String = "'VBASHADOWSOURE-INSERTED-END"


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
        If tempComponent.Name = OWNMODULENAME Then GoTo NEXTLOOP
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
        If tempComponent.Name = OWNMODULENAME Then GoTo NEXTLOOP
        
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

'--
' アドインファイルの出力
'--
Public Sub exportAdinFile()
      Dim tempFilePath As String
      tempFilePath = ThisWorkbook.Path & "\" & "TEMP_" & ThisWorkbook.Name
      ThisWorkbook.SaveCopyAs tempFilePath
      
      Dim tempbook As Workbook: Set tempbook = Workbooks.Open(tempFilePath)
      
      Dim component As Object
      For Each component In tempbook.VBProject.VBComponents
          'アドイン化した際にシート・ブックVBAは消えるため、気にせずクリアする。
          If component.Name = "ThisWorokbook" Then
            component.CodeModule.DeleteLines 1, component.CodeModule.CountOfLines
          ElseIf component.Name = OWNMODULENAME Then
            Call tempbook.VBProject.VBComponents.Remove(component)
          End If
      Next
      
      tempbook.SaveAs _
        Filename:=ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".", "_") & "_src" & "\addin.xlam", _
        FileFormat:=xlOpenXMLAddIn
      tempbook.Close
End Sub

'--
' 初期化コードの削除
'--
Sub deleteOwnCode()
      Dim component As VBComponent
      For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Name = "ThisWorkbook" Then
          
          Dim currentLine As Long
          Dim isFoundStartComment As Boolean
          Dim isFoundEndComment As Boolean
          
          While True
            Dim startLine As Long: startLine = currentLine
            isFoundStartComment = component.CodeModule.Find(InsertedStartSymbol, startLine, 0, 0, 0)
            Dim endLine As Long: endLine = startLine + 1
            isFoundEndComment = component.CodeModule.Find(InsertedEndSymbol, endLine, 0, 0, 0)
                
            If Not isFoundStartComment Or Not isFoundEndComment Then
              GoTo LOOP_BREAK
            End If
            component.CodeModule.DeleteLines startLine, endLine - startLine + 1
            currentLine = endLine + 1
          Wend
LOOP_BREAK:

        End If
      Next
End Sub


'--
' 初期化
'--
Private Sub initVBASHADOWSOURCE()
    Dim code As String
    Const INSERTEDSYMBOL = "'VBASHADOWSOURE-INSERTED"
    code = InsertedStartSymbol & vbCrLf _
         & "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)" & vbCrLf _
         & "  'コードをエクスポート" & vbCrLf _
         & "  Call exportVBAFiles" & vbCrLf _
         & "End Sub" & vbCrLf _
         & "Private Sub Workbook_Open()" & vbCrLf _
         & " 'コードをインポート" & vbCrLf _
         & "  Call importVBAFiles" & vbCrLf _
         & "End Sub" & vbCrLf _
         & InsertedEndSymbol
    
    Dim tempComponent As Object
    For Each tempComponent In ThisWorkbook.VBProject.VBComponents
        If tempComponent.Name = "ThisWorkbook" Then
            If Not tempComponent.CodeModule.Find(InsertedStartSymbol, 0, 0, 0, 0) Then
                tempComponent.CodeModule.InsertLines 1, code
            End If
        End If
    Next
End Sub

