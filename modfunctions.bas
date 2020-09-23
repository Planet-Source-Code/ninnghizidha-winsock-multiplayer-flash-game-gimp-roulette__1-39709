Attribute VB_Name = "modFunctions"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
    As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpRetunedString As String, ByVal nSize As Long, _
    ByVal lpFilename As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long


Public Function SetOption(iAppName As String, iKeyName As String, iKeyString As String)
Dim ret
        ret = WritePrivateProfileString(iAppName, iKeyName, iKeyString, App.Path & "\gimp.ini")
End Function

Public Function GetOption(iAppName As String, iKeyName As String) As String
    Dim iStr As String
    iStr = String(255, Chr(0))
    GetOption = Left(iStr, GetPrivateProfileString(iAppName, ByVal iKeyName, "", iStr, Len(iStr), App.Path & "\gimp.ini"))
End Function

Public Sub genPercentBar(pPictureBox As PictureBox, pLabelPercent As Label, plngMax As Long, plngPart As Long)
Dim intPercentValue As Integer
    If plngMax > 0 Then
        intPercentValue = plngPart / plngMax * 1000
        intPercentValue = Round(intPercentValue)
        
        pPictureBox.Cls
        pPictureBox.ScaleMode = 0
        pPictureBox.ScaleWidth = 1000
        pPictureBox.ScaleHeight = 10
        pLabelPercent.Caption = Round(intPercentValue / 10) & "%"
        
        pPictureBox.Line (0, 0)-(intPercentValue, pPictureBox.ScaleHeight), pPictureBox.FillColor, BF
    End If
End Sub

