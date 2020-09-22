Attribute VB_Name = "modFileVersion"
Option Explicit
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
End Type

   Dim rc                As Long
   Dim lDummy            As Long
   Dim sBuffer()         As Byte
   Dim lBufferLen        As Long
   Dim lVerPointer       As Long
   Dim udtVerBuffer      As VS_FIXEDFILEINFO
   Dim lVerbufferLen     As Long

Public FileVersion As String

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

Public Function FileVersionChk(FullFileName As String) As String
 
 '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      Exit Function
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
   '**** Determine File Version number ****

   FileVersionChk = Format$(udtVerBuffer.dwFileVersionMSh) & Format$(udtVerBuffer.dwFileVersionMSl) & Format$(udtVerBuffer.dwFileVersionLSh) & Format$(udtVerBuffer.dwFileVersionLSl)
  
End Function

