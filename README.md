# REST_API_Excel_VBA_MacOS_64bit
Option Explicit
Private Declare PtrSafe Function popen Lib "/usr/lib/libc.dylib" (ByVal Command As String, ByVal Mode As String) As LongPtr
Private Declare PtrSafe Function pclose Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As Long
Private Declare PtrSafe Function fread Lib "/usr/lib/libc.dylib" (ByVal outStr As String, ByVal size As LongPtr, ByVal items As LongPtr, ByVal stream As LongPtr) As Long
Private Declare PtrSafe Function feof Lib "/usr/lib/libc.dylib" (ByVal file As LongPtr) As LongPtr

Function execShell(Command As String, Optional ByRef exitCode As Long) As String
    Dim file As LongPtr
    file = popen(Command, "r")

    If file = 0 Then
        Exit Function
    End If

    While feof(file) = 0
        Dim chunk As String
        Dim read As Long
        chunk = Space(500)
        read = fread(chunk, 1, Len(chunk) - 1, file)
        If read > 0 Then
            chunk = Left$(chunk, read)
            execShell = execShell & chunk
        End If
    Wend

    exitCode = pclose(file)
    
End Function



Function HTTPGet(sUrl As String, sQuery As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl -X GET " & sQuery & "" & " " & sUrl
    sResult = execShell(sCmd, lExitCode)
    HTTPGet = sResult

End Function

Function HTTPPost(sUrl As String, sQuery1 As String, sQuery2 As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl -X POST " & sQuery1 & "" & " -d " & sQuery2 & "" & " " & sUrl
    sResult = execShell(sCmd, lExitCode)
    HTTPPost = sResult

End Function

Function HTTPPut(sUrl As String, sQuery1 As String, sQuery2 As String) As String

    Dim sCmd As String
    Dim sResult As String
    Dim lExitCode As Long

    sCmd = "curl -X PUT " & sQuery1 & "" & " -d " & sQuery2 & "" & " " & sUrl
    sResult = execShell(sCmd, lExitCode)
    HTTPPut = sResult

End Function




'GET-запросы
Sub SendGETRequest()

Dim i As Integer
Dim j As Integer
Dim result As String
Dim URL As String
Dim Auth As String

a = Timer

    'Подсчет заполненных ячеек первого столбца
    i = 1
        Do While Not IsEmpty(Cells(i, 1))
        i = i + 1
        Loop
    i = i - 1

    'Цикл, который отправляет запрос от 2 до последнего элемента
    For j = 2 To i
        URL = Range("I" & j)
        Auth = Range("H" & j)
        result = HTTPGet(URL, Auth)
        Range("J" & j).Value = result
        'Application.Wait (Now + TimeValue("0:00:01"))
    Next j

MsgBox Timer - a

End Sub


'PUT-запросы
Sub SendPUTRequest()

Dim i As Integer
Dim j As Integer
Dim result As String
Dim URL As String
Dim Auth As String
Dim Message As String


a = Timer

    'Подсчет заполненных ячеек первого столбца
    i = 1
        Do While Not IsEmpty(Cells(i, 1))
        i = i + 1
        Loop
    i = i - 1

    'Цикл, который отправляет запрос от 2 до последнего элемента
    For j = 2 To i
        Message = Range("D" & j)
        URL = Range("I" & j)
        Auth = Range("H" & j)
        result = HTTPPut(URL, Auth, Message)
        Range("J" & j).Value = result
        'Application.Wait (Now + TimeValue("0:00:01"))
    Next j

MsgBox Timer - a

End Sub


'POST-запросы
Sub SendPOSTRequest()

Dim i As Integer
Dim j As Integer
Dim result As String
Dim URL As String
Dim Auth As String
Dim Message As String

a = Timer

    'Подсчет заполненных ячеек первого столбца
    i = 1
        Do While Not IsEmpty(Cells(i, 1))
        i = i + 1
        Loop
    i = i - 1

    'Цикл, который отправляет запрос от 2 до последнего элемента
    For j = 2 To i
        Message = Range("D" & j)
        URL = Range("I" & j)
        Auth = Range("H" & j)
        result = HTTPPost(URL, Auth, Message)
        Range("J" & j).Value = result
        'Application.Wait (Now + TimeValue("0:00:01"))
    Next j

MsgBox Timer – a
