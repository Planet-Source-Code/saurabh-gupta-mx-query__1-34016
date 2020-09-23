VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MX Query Demo"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   1320
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Mail Server"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Preference"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4080
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
      Height          =   255
      Left            =   1133
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "hotmail.com"
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblResult 
      Caption         =   "Results :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Domain :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "DNS server to use :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DNSrecieved As Boolean
Private dnsReply() As Byte
Private time2 As Long

Private Sub Command1_Click()
    ListView1.ListItems.Clear
    Text1.Text = Trim(Text1.Text)
    Text2.Text = Trim(Text2.Text)
    If InStr(Text1.Text, ".") = 0 Then
        MsgBox "Please enter a correct DNS server IP"
        Exit Sub
    End If
    If InStr(Text2.Text, ".") = 0 Then
        MsgBox "Please enter a correct Domain"
        Exit Sub
    End If
    Dim bestPref As String
    Command1.Enabled = False
    bestPref = MX_Query(Text1.Text, Text2.Text)
    If bestPref = "" Then
        MsgBox "No MX records found. Please check that the DNS server is correct and the domain exists", vbExclamation
    Else
        MsgBox CStr(ListView1.ListItems.Count) + " MX entries found from DNS server " + Text1.Text + vbCrLf + _
               "Prefered mail exchange server : " + bestPref, vbInformation
    End If
    Command1.Enabled = True
End Sub

Private Sub Form_Load()
    Text1.Text = GetDNSinfo
    
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print Description
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    DNSrecieved = True
    ReDim dnsReply(bytesTotal) As Byte
    Winsock2.GetData dnsReply, vbArray + vbByte
End Sub
Private Sub Timer2_Timer()
    time2 = time2 + 1
    If time2 > 10 Then Timer2.Enabled = False
End Sub

Private Function MX_Query(DNS_Addr As String, ByVal Domain_Addr As String) As String
    Dim IpAddr As Long
    Dim iRC As Integer
    Dim dnsHead As DNS_HEADER
    Dim iSock As Integer
    
    ' Set the DNS parameters
    dnsHead.qryID = htons(&H11DF)
    dnsHead.options = DNS_RECURSION
    dnsHead.qdcount = htons(1)
    dnsHead.ancount = 0
    dnsHead.nscount = 0
    dnsHead.arcount = 0
    
    ' Query Variables
    Dim dnsQuery() As Byte
    Dim sQName As String
    Dim dnsQueryNdx As Integer
    Dim iTemp As Integer
    Dim iNdx As Integer
    dnsQueryNdx = 0
    ReDim dnsQuery(4000)
    
    ' Setup the dns structure to send the query in
    
    ' First goes the DNS header information
    MemCopy dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    ' Then the domain name (as a QNAME)
    sQName = MakeQName(Domain_Addr)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
    Wend

    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    ' Null terminate the string
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    ' The type of query (15 means MX query)
    iTemp = htons(15)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ' The class of query (1 means INET)
    iTemp = htons(1)
    MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    On Error Resume Next
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    ' Send the query to the DNS server
    Winsock2.RemoteHost = DNS_Addr
    Winsock2.RemotePort = 53
    DNSrecieved = False
    Winsock2.SendData dnsQuery
    'Err.Clear
    
    Timer2.Enabled = True
    time2 = 0
    Do While Not DNSrecieved And time2 < 10
        If Winsock2.State = sckError Then
            Timer2.Enabled = False
            Exit Function
        End If
        DoEvents
    Loop
    Timer2.Enabled = False
    If Not DNSrecieved Then Exit Function
    
    
    Dim iAnCount As Integer
    ' Get the number of answers
    MemCopy iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    ' Parse the answer buffer
    MX_Query = Trim(GetMXName(dnsReply(), 12, iAnCount))
End Function


