Attribute VB_Name = "Module1"
'here are all the declarations of the api and the types are used from the apis
Private Declare Function RasGetCredentials Lib "rasapi32.dll" Alias "RasGetCredentialsA" _
 (ByVal lpcstr As String, ByVal lpcstr As String, ByRef TLPRASCREDENTIALSA As RASCREDENTIALS) _
 As Long
Private Type RASCREDENTIALS
    dwSize As Long
    dwMask As Long
    szUserName As String
    szPassword As String
    szDomain As String
End Type
Public Type RASENTRYNAME95
    dwSize As Long
    szEntryname(256) As Byte
End Type

Public Declare Function RasEnumEntriesA Lib "rasapi32.dll" _
    (ByVal reserved As String, ByVal lpszPhonebook As String, _
    lprasentryname As Any, lpcb As Long, lpcEntries As Long) _
    As Long




Public Declare Function RasGetEntryDialParams _
      Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, _
        lpRasDialParams As Any, _
        blnPasswordRetrieved As Long) As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)


Public Type VBRasDialParams
    EntryName As String
    PhoneNumber As String
    CallbackNumber As String
    UserName As String
    Password As String
    Domain As String
    SubEntryIndex As Long
    RasDialFunc2CallbackId As Long
End Type
'VB "friendly" RasDialParams Structure



Function VBRasSetEntryDialParams _
              (strPhonebook As String, bytesIn() As Byte, _
               blnRemovePassword As Boolean) As Long
   
 '  VBRasSetEntryDialParams = RasSetEntryDialParams _
  '             (strPhonebook, bytesIn(0), blnRemovePassword)
End Function
Sub CopyByteToTrimmedString(strToCopyTo As String, _
                              bPos As Byte, lngMaxLen As Long)
   Dim strTemp As String, lngLen As Long
   strTemp = String(lngMaxLen + 1, 0)
   CopyMemory ByVal strTemp, bPos, lngMaxLen
   lngLen = InStr(strTemp, Chr$(0)) - 1
   strToCopyTo = Left$(strTemp, lngLen)
End Sub

Sub CopyStringToByte(bPos As Byte, _
                        strToCopy As String, lngMaxLen As Long)
   Dim lngLen As Long
   lngLen = Len(strToCopy)
   If lngLen = 0 Then
      Exit Sub
   ElseIf lngLen > lngMaxLen Then
      lngLen = lngMaxLen
   End If
   CopyMemory bPos, ByVal strToCopy, lngLen
End Sub
Function BytesToVBRasDialParams(bytesIn() As Byte, _
            udtVBRasDialParamsOUT As VBRasDialParams) As Boolean
   
   Dim iPos As Long, lngLen As Long
   Dim dwSize As Long
   On Error GoTo badBytes
   
   CopyMemory dwSize, bytesIn(0), 4
   
   If dwSize = 816& Then
      lngLen = 21&
   ElseIf dwSize = 1060& Or dwSize = 1052& Then
      lngLen = 257&
   Else
      'unkown size
      Exit Function
   End If
   iPos = 4
   With udtVBRasDialParamsOUT
      CopyByteToTrimmedString .EntryName, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyByteToTrimmedString .PhoneNumber, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyByteToTrimmedString .CallbackNumber, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyByteToTrimmedString .UserName, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyByteToTrimmedString .Password, bytesIn(iPos), lngLen
      iPos = iPos + lngLen: lngLen = 16
      CopyByteToTrimmedString .Domain, bytesIn(iPos), lngLen
      
      If dwSize > 1052& Then
         CopyMemory .SubEntryIndex, bytesIn(1052), 4&
         CopyMemory .RasDialFunc2CallbackId, bytesIn(1056), 4&
      End If
   End With
   BytesToVBRasDialParams = True
   Exit Function
badBytes:
   'error handling goes here ??
   BytesToVBRasDialParams = False
End Function



Public Sub DUN_Services(DUN_Array() As String)
'Pass in Empty array for DUN_Array
    Dim s As Long, ln As Long, conname As String, i As Long
    Dim r(255) As RASENTRYNAME95
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    Call RasEnumEntriesA(vbNullString, vbNullString, r(0), s, ln)
    ln = ln - 1
    ReDim DUN_Array(ln)
    For i = 0 To ln
        conname = StrConv(r(i).szEntryname(), vbUnicode)
        
        DUN_Array(i) = Left$(conname, InStr(conname, _
          vbNullChar) - 1)
         'RasGetEntryDialParams
    Next i
End Sub
Function VBRasGetEntryDialParams _
              (bytesOut() As Byte, _
          strPhonebook As String, strEntryName As String, _
               Optional blnPasswordRetrieved As Boolean) As Long
 
   Dim rtn As Long
     Dim blnPsswrd As Long
   Dim bLens As Variant
   Dim lngLen As Long, i As Long
   bLens = Array(1060&, 1052&, 816&)
   'try our three different sizes for RasDialParams
   'eatch OS version has is own structure size
   For i = 0 To 2
      lngLen = bLens(i)
      ReDim bytesOut(lngLen - 1)
      CopyMemory bytesOut(0), lngLen, 4
      If lngLen = 816& Then
         CopyStringToByte bytesOut(4), strEntryName, 20
      Else
         CopyStringToByte bytesOut(4), strEntryName, 256
      End If
      rtn = RasGetEntryDialParams(strPhonebook, bytesOut(0), blnPsswrd)
    
  
      If rtn = 0 Then Exit For
   Next i
   
   blnPasswordRetrieved = blnPsswrd
   VBRasGetEntryDialParams = rtn
End Function


