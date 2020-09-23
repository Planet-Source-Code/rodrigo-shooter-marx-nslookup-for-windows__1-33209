Attribute VB_Name = "getip"
Option Explicit ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSADATA) As Long
  
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function inet_addr Lib "wsock32" _
  (ByVal s As String) As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
Public Const IP_SUCCESS As Long = 0
Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1

Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
Private Const AF_INET As Long = 2

Public Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32" _
  (ByVal hostname As String) As Long
  


Private Declare Function lstrlenA Lib "kernel32" _
  (lpString As Any) As Long


    



Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim success As Long
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub


Public Function GetIPFromHostName(ByVal sHostName As String) As String

  'converts a host name to an IP address.

   Dim nbytes As Long
   Dim ptrHosent As Long
   Dim ptrName As Long
   Dim ptrAddress As Long
   Dim ptrIPAddress As Long
   Dim sAddress As String
   
   sAddress = Space$(4)

   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

     'assign pointer addresses and offset
     
     'The Address is offset 12 bytes from the start of
     'the HOSENT structure. Note: Here we are retrieving
     'only the first address returned. To return more than
     'one, define sAddress as a string array and loop through
     'the 4-byte ptrIPAddress members returned. The last
     'item is a terminating null. All addresses are returned
     'in network byte order.
      ptrAddress = ptrHosent + 12
      
     'get the IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4

      GetIPFromHostName = IPToText(sAddress)

   End If
   
End Function


Private Function IPToText(ByVal IPAddress As String) As String

   IPToText = CStr(Asc(IPAddress)) & "." & _
              CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
              CStr(Asc(Mid$(IPAddress, 4, 1)))
              
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





  
  






Public Function GetHostNameFromIP(ByVal sAddress As String) As String

   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
   
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               GetHostNameFromIP = sAddress
            End If
         
         Else: MsgBox "Call to gethostbyaddr failed."
         End If 'If ptrHosent
      
      SocketsCleanup
      
      Else: MsgBox "String passed is an invalid IP."
      End If 'If hAddress
   
   Else: MsgBox "Sockets failed to initialize."
   End If  'If SocketsInitialize
      
End Function
