


'===================== Funções para conversão de charset

Function ConvertFileCharset(ByVal oFile As Object) As String

    On Error GoTo ErrChk
    
   Const adTypeBinary = 1
   Const adTypeText = 2
   Const bOverwrite = True
   Const bAsASCII = False
 
   Dim oFS: Set oFS = CreateObject("Scripting.FileSystemObject")
 
   Dim oFrom: Set oFrom = CreateObject("ADODB.Stream")
   Dim sFrom: sFrom = "Windows-1252"
   Dim sFFSpec: sFFSpec = oFile.path
   Dim oTo: Set oTo = CreateObject("ADODB.Stream")
   Dim sTo: sTo = "utf-8"
   Dim sTFSpec: sTFSpec = oFile.ParentFolder.path & "\utf" & oFile.Name
 
   oFrom.Type = adTypeText
   oFrom.Charset = sFrom
   oFrom.Open
   oFrom.LoadFromFile sFFSpec
   strFullText = oFrom.ReadText
   oFrom.Close
 
   oTo.Type = adTypeText
   oTo.Charset = sTo
   oTo.Open
   oTo.WriteText strFullText
   'WScript.Echo oTo.Size & " Bytes in " & sTFSpec
   oTo.SaveToFile sTFSpec
   oTo.Close
   
   ConvertFileCharset = sTFSpec
   
   Exit Function
   
ErrChk:
    MsgBox Err.Description
    Stop
    Resume

End Function

Public Function ReturnCharset(ByVal filePath As String, Optional verifyANSI As Boolean = True) As Integer
    Const bytByte0Unicode_c As Byte = 255
    Const bytByte1Unicode_c As Byte = 254
    Const bytByte0UnicodeBigEndian_c As Byte = 254
    Const bytByte1UnicodeBigEndian_c As Byte = 255
    Const bytByte0UTF8_c As Byte = 239
    Const bytByte1UTF8_c As Byte = 187
    Const bytByte2UTF8_c As Byte = 191
    Const lngByte0 As Long = 0
    Const lngByte1 As Long = 1
    Const lngByte2 As Long = 2
    Dim bytHeader() As Byte
    Dim eRtnVal As abCharsets
    On Error GoTo Err_Hnd
    bytHeader() = GetFileBytes(filePath, lngByte2)
    Select Case bytHeader(lngByte0)
    Case bytByte0Unicode_c
        If bytHeader(lngByte1) = bytByte1Unicode_c Then
            eRtnVal = abCharsets.abUnicode
        End If
    Case bytByte0UnicodeBigEndian_c
        If bytHeader(lngByte1) = bytByte1UnicodeBigEndian_c Then
            eRtnVal = abCharsets.abUnicodeBigEndian
        End If
    Case bytByte0UTF8_c
        If bytHeader(lngByte1) = bytByte1UTF8_c Then
            If bytHeader(lngByte2) = bytByte2UTF8_c Then
                eRtnVal = abCharsets.abUTF8
            End If
        End If
    End Select
    If Not CBool(eRtnVal) Then
        If verifyANSI Then
            If IsANSI(filePath) Then
                eRtnVal = abCharsets.abANSI
            Else
                eRtnVal = abCharsets.ebUnknown
            End If
        Else
            eRtnVal = abCharsets.abANSI
        End If
    End If
Exit_Proc:
    On Error Resume Next
    Erase bytHeader
    ReturnCharset = eRtnVal
    Exit Function
Err_Hnd:
    eRtnVal = abCharsets.abError
    Resume Exit_Proc
End Function

Private Function IsANSI(ByVal filePath As String) As Boolean
    Const lngKeyCodeNullChar_c As Long = 0
    Dim bytFile() As Byte
    Dim lngIndx As Long
    Dim lngUprBnd As Long
    bytFile = GetFileBytes(filePath)
    lngUprBnd = UBound(bytFile)
    For lngIndx = 0 To lngUprBnd
        If bytFile(lngIndx) = lngKeyCodeNullChar_c Then
            Exit For
        End If
    Next
    Erase bytFile
    IsANSI = (lngIndx > lngUprBnd)
End Function

Public Function GetFileBytes(ByVal path As String, Optional ByVal truncateToByte As Long = -1) As Byte()
    Dim lngFileNum As Long
    Dim bytRtnVal() As Byte
    If truncateToByte < 0 Then
        truncateToByte = FileLen(path) - 1
    End If
    lngFileNum = FreeFile
    If FileExists(path) Then
        Open path For Binary Access Read As lngFileNum
        ReDim bytRtnVal(truncateToByte) As Byte
        Get lngFileNum, , bytRtnVal
        Close lngFileNum
    End If
    GetFileBytes = bytRtnVal
    Erase bytRtnVal
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    FileExists = CBool(LenB(Dir(filePath, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)))
End Function

Public Function CharsetToString(ByVal value As abCharsets) As String
    Dim strRtnVal As String
    Select Case value
    Case abCharsets.abANSI
        strRtnVal = "us-ascii"
    Case abCharsets.abUTF8
        strRtnVal = "utf-8"
    Case Else
        strRtnVal = "Unicode"
    End Select
    CharsetToString = strRtnVal
End Function

'===================== Fim Funções para conversão de charset

Sub ChangeCSVCharacter(ByVal filePath As String)

    On Error GoTo ErrChk

    sTemp = ""
    Open filePath For Input As #1
    Do Until EOF(1)
        Line Input #1, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop
    Close #1
    
    sTemp = Replace(Replace(sTemp, ",", "."), ";", ",")
    
    Open filePath For Output As #1
    Print #1, sTemp
    Close #1
    
    Exit Sub

ErrChk:
    MsgBox Err.Description
    Stop
    Resume

End Sub
