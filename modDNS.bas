Attribute VB_Name = "modDNS"
Option Explicit


Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 1
Public Const MIB_IF_TYPE_TOKENRING = 2
Public Const MIB_IF_TYPE_FDDI = 3
Public Const MIB_IF_TYPE_PPP = 4
Public Const MIB_IF_TYPE_LOOPBACK = 5
Public Const MIB_IF_TYPE_SLIP = 6

Type IP_ADDR_STRING
            Next As Long
            IpAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type
Type FIXED_INFO
            HostName As String * MAX_HOSTNAME_LEN
            DomainName As String * MAX_DOMAIN_NAME_LEN
            CurrentDnsServer As Long
            DnsServerList As IP_ADDR_STRING
            NodeType As Long
            ScopeId  As String * MAX_SCOPE_ID_LEN
            EnableRouting As Long
            EnableProxy As Long
            EnableDns As Long
End Type

Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
'Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Public Const DNS_RECURSION As Byte = 1



Public Type DNS_HEADER
    qryID As Integer
    options As Byte
    response As Byte
    qdcount As Integer
    ancount As Integer
    nscount As Integer
    arcount As Integer
End Type

Public Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const hostent_size = 16

Public Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Public fMainForm As Form1

Sub main()
    Set fMainForm = New Form1
    fMainForm.Show
End Sub

Public Function MakeQName(sDomain As String) As String
    Dim iQCount As Integer      ' Character count (between dots)
    Dim iNdx As Integer         ' Index into sDomain string
    Dim iCount As Integer       ' Total chars in sDomain string
    Dim sQName As String        ' QNAME string
    Dim sDotName As String      ' Temp string for chars between dots
    Dim sChar As String         ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName
End Function

Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    Dim iCompress As Integer        ' Compression index (index into original buffer)
    Dim iChCount As Integer         ' Character count (number of chars to read from buffer)
        
    ' While we didn't encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
End Sub


Public Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    Dim iChCount As Integer     ' Character counter
    Dim sTemp As String         ' Holds original query string
    
    Dim iMXLen As Integer
    Dim iBestPref As Integer    ' Holds the "best" preference number (lowest)
    Dim sBestMX As String       ' Holds the "best" MX record (the one with the lowest preference)
    
    iBestPref = -1
    
    ParseName dnsReply(), iNdx, sTemp
    ' Step over null
    iNdx = iNdx + 2
    
    ' Step over 6 bytes (not sure what the 6 bytes are, but all other
    '   documentation shows steping over these 6 bytes)
    iNdx = iNdx + 6
    
    Dim xItem As ListItem
    
    On Error Resume Next
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6
            
            ' Read the MX data length specifier
            '              (not needed, hence why it's commented out)
            MemCopy iMXLen, dnsReply(iNdx), 2
            iMXLen = ntohs(iMXLen)
            
            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            MemCopy iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            Dim iNdx2 As Integer
            iNdx2 = iNdx
            ParseName dnsReply(), iNdx2, sName
            If (iBestPref = -1 Or iPref < iBestPref) Then
                iBestPref = iPref
                sBestMX = sName
            End If
            Set xItem = fMainForm.ListView1.ListItems.Add(Text:=sName)
            xItem.ListSubItems.Add Text:=iPref
            
            iNdx = iNdx + iMXLen + 1
            ' Step over 3 useless bytes
            'iNdx = iNdx + 3
        Else
            GetMXName = sBestMX
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    GetMXName = sBestMX
End Function

Public Function GetDNSinfo() As String
    Dim error As Long
    Dim FixedInfoSize As Long
    Dim strDNS  As String
    Dim FixedInfo As FIXED_INFO
    Dim Buffer As IP_ADDR_STRING
    Dim FixedInfoBuffer() As Byte
    
    FixedInfoSize = 0
    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetNetworkParams sizing failed with error: " & error
           Exit Function
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)
    

    error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    If error = 0 Then
        CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
        strDNS = FixedInfo.DnsServerList.IpAddress
        strDNS = Replace(strDNS, vbCr, "")
        strDNS = Replace(strDNS, vbLf, "")
        strDNS = Replace(strDNS, vbNullChar, "")
        strDNS = Trim(strDNS)
        GetDNSinfo = strDNS
    End If
        
End Function

