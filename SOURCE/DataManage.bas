Attribute VB_Name = "modDataFunctions"
Option Explicit

Public Declare Function WritePrivateProfileString& Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

    


Public Sub WriteEoC(EoCSection As String, EoCKey As String, EoCValue As String, EoCFile As String)
    Call WritePrivateProfileString(EoCSection, EoCKey, EoCValue, EoCFile)
End Sub

Public Function ReadEoC(Section As String, KeyName As String, FileName As String, Default As String) As String
    Dim sRet As String

    sRet = String(255, Chr(0))

    ReadEoC = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, Default, sRet, Len(sRet), FileName))
End Function

Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    On Error GoTo GetVar_Error

    szReturn = vbNullString

    sSpaces = Space(5000)

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    On Error GoTo 0
    Exit Function

GetVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetVar of Module modDatabase"
End Function

Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error

    Call WritePrivateProfileString(Header, Var, Value, file)

    On Error GoTo 0
    Exit Sub

PutVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub Main()

Call SystemFileChecker

frmSTART.Visible = True

End Sub

Public Function EventExist(EventName As String) As Boolean


If GetVar(App.Path & "\Characters\" & CName & "Events" & ".EoC", "Events", EventName) = "" Then
 EventExist = False
  
  Else
   EventExist = True
    End If
    

End Function
 

