Attribute VB_Name = "DartFs"
Option Explicit

Public Sub UpdateData()
    Dim py As String, sc As String, cmd As String
    On Error GoTo Fail
    py = ThisWorkbook.Names("cfg_python").RefersToRange.Value
    sc = ThisWorkbook.Names("cfg_script").RefersToRange.Value
    ThisWorkbook.Save
    ThisWorkbook.Names("cfg_status").RefersToRange.Value = "업데이트 실행 중... " & Format(Now, "hh:nn:ss")
    cmd = """" & py & """ """ & sc & """ --workbook """ & ThisWorkbook.FullName & """"
    Shell cmd, vbMinimizedNoFocus
    Exit Sub
Fail:
    ThisWorkbook.Names("cfg_status").RefersToRange.Value = "실행 실패: " & Err.Description
End Sub
