VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Enumerator (Beta)"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "NetServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3420
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7752
            Text            =   "System messages will display here"
            TextSave        =   "System messages will display here"
            Key             =   "msg"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Errors, Messages or Additional System Information "
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "1:27 PM"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "System Time"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "1/05/2003"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "System Date"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Send your message"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "NetServer.frx":0442
      Left            =   480
      List            =   "NetServer.frx":044C
      TabIndex        =   2
      ToolTipText     =   "Highlight a system to send a message (NT systems only)"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "NetServer.frx":048F
      Left            =   480
      List            =   "NetServer.frx":04D9
      TabIndex        =   1
      ToolTipText     =   "Select the System Type you wish to view, then click below "
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Type your message here (NT systems only)"
      Top             =   720
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Browse List Information"
      Height          =   3135
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton Command3 
         Caption         =   "&About"
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "NetServer.frx":06B9
         Left            =   240
         List            =   "NetServer.frx":06BB
         TabIndex        =   0
         ToolTipText     =   "Select a Domain/Workgroup to browse"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Domain/Workgroup to Enumerate:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Send a Message:"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "System Types to Enumerate:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    GetDomain (SV_TYPE_DOMAIN_ENUM)
    
End Sub

Private Sub Combo1_LostFocus()
    
    If Combo1 = "LAN Manager Workstations" Then
        Call ListServers(SV_TYPE_WORKSTATION)
    End If
    
    If Combo1 = "LAN Manager Servers" Then
        Call ListServers(SV_TYPE_SERVER)
    End If
    
    If Combo1 = "SQL Servers" Then
        Call ListServers(SV_TYPE_SQLSERVER)
    End If
    
    If Combo1 = "Primary Domain Controllers" Then
        Call ListServers(SV_TYPE_DOMAIN_CTRL)
    End If
    
    If Combo1 = "Backup Domain Controllers" Then
        Call ListServers(SV_TYPE_DOMAIN_BAKCTRL)
    End If
    
    If Combo1 = "All Domain Controllers" Then
        Call ListServers(SV_TYPE_DC)
    End If
    
    If Combo1 = "Timesource Servers" Then
        Call ListServers(SV_TYPE_TIMESOURCE)
    End If
    
    If Combo1 = "Apple File Protocol Servers" Then
        Call ListServers(SV_TYPE_AFP)
    End If
    
    'If Combo1 = "Novell Servers" Then
    '    Call ListServers(SV_TYPE_NOVELL)
    'End If
    
    If Combo1 = "LM 2.x Domain Members" Then
        Call ListServers(SV_TYPE_DOMAIN_MEMBER)
    End If
    
    If Combo1 = "Local Browse List (MB Only)" Then
        Call ListServers(SV_TYPE_LOCAL_LIST_ONLY)
    End If
    
    If Combo1 = "Print Servers" Then
        Call ListServers(SV_TYPE_PRINT)
    End If
    
    If Combo1 = "Dial-in Servers" Then
        Call ListServers(SV_TYPE_DIALIN)
    End If
    
    'If Combo1 = "Xenix Servers" Then
    '   Call ListServers(SV_TYPE_XENIX_SERVER)
    'End If
    
    If Combo1 = "MS Novell File & Print Servers" Then
        Call ListServers(SV_TYPE_MFPN)
    End If
    
    If Combo1 = "Windows NT (S&W)" Then
        Call ListServers(SV_TYPE_NT)
    End If
    
    If Combo1 = "WfW Servers" Then
        Call ListServers(SV_TYPE_WFW)
    End If
    
    'If Combo1 = "Non-DC NT Servers" Then
    '    Call ListServers(SV_TYPE_SERVER_NT)
    'End If
    
    If Combo1 = "Potential Master Browsers" Then
        Call ListServers(SV_TYPE_POTENTIAL_BROWSER)
    End If
    
    If Combo1 = "Backup Master Browsers" Then
        Call ListServers(SV_TYPE_BACKUP_BROWSER)
    End If
    
    If Combo1 = "Master Browser Servers" Then
        Call ListServers(SV_TYPE_MASTER_BROWSER)
    End If
    
    If Combo1 = "Domain Master Browsers" Then
        Call ListServers(SV_TYPE_DOMAIN_MASTER)
    End If
    
    If Combo1 = "Windows 95 and Later" Then
        Call ListServers(SV_TYPE_WINDOWS)
    End If
    
    If Combo1 = "All Server Types" Then
        Call ListServers(SV_TYPE_ALL)
    End If
    
End Sub

Private Sub Command1_Click()
    Dim lReturnCode As Long
    Dim sUnicodeToName As String
    Dim sUnicodeFromName As String
    Dim sUnicodeMessage As String
    Dim lMessageLength As Long
    
    sUnicodeFromName = StrConv(GetLocalSystemName, vbUnicode)
    sUnicodeToName = StrConv(List1.Text, vbUnicode)
    sUnicodeMessage = StrConv(Text1.Text, vbUnicode)
    lMessageLength = Len(sUnicodeMessage)
    
    ' Hourglass pointer
    MousePointer = vbHourglass
    StatusBar1.Panels("msg").Text = vbNullString
    
    ' Send the message
    lReturnCode = NetMessageBufferSend("", sUnicodeToName, sUnicodeFromName, sUnicodeMessage, lMessageLength)
    
    ' Prove some feedback about the send action
    If lReturnCode = 0 Then
        Text1.Text = vbNullString
        StatusBar1.Panels("msg").Text = "Your message was successfully sent... will they read it?"
    Else
        StatusBar1.Panels("msg").Text = "Error - Return code: " & CStr(lReturnCode)
    End If
    
    ' Default pointer
    MousePointer = vbDefault
    
End Sub

Private Sub Command2_Click()
    
    Unload Me
    End
    
End Sub

Private Sub Command3_Click()
    
    frmAbout.Show
    
End Sub

Private Sub GetDomain(lType As Long)
    Dim lReturn As Long
    Dim Server_Info As Long
    Dim lEntries As Long
    Dim lTotal As Long
    Dim lMax As Long
    Dim vResume As Variant
    Dim tServer_info_101 As SERVER_INFO_101
    Dim sServer As String
    Dim sDomain As String
    Dim lServerInfo101StructPtr As Long
    Dim X As Long, i As Long
    Dim bBuffer(512) As Byte
    
    Form1.Combo2.Clear
    
    lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
    
    If lReturn <> 0 Then
        StatusBar1.Panels("msg").Text = "Error " + Str$(lReturn) + " when trying to obtain server list " + Str$(lTotal)
        Exit Sub
    End If
    
    X = 1
    lServerInfo101StructPtr = Server_Info
    
    Do While X <= lTotal
        
        RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
        
        lstrcpyW bBuffer(0), tServer_info_101.ptr_name
        
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & Chr$(bBuffer(i))
            i = i + 2
        Loop
        Form1.Combo2.AddItem sServer
        DoEvents
        X = X + 1
        sServer = ""
        lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
            
    Loop
            
    lReturn = NetApiBufferFree(Server_Info)
            
End Sub

Private Sub ListServers(lType As Long)
    Dim lReturn As Long
    Dim Server_Info As Long
    Dim lEntries As Long
    Dim lTotal As Long
    Dim lMax As Long
    Dim vResume As Variant
    Dim tServer_info_101 As SERVER_INFO_101
    Dim sServer As String
    Dim sDomain As String
    Dim lServerInfo101StructPtr As Long
    Dim X As Long, i As Long
    Dim bBuffer(512) As Byte
    
    sDomain = StrConv(Combo2.Text, vbUnicode)
    
    Form1.List1.Clear
    
    lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
    
    If lReturn <> 0 Then
        StatusBar1.Panels("msg").Text = "Error " + Str$(lReturn) + " when trying to obtain server list " + Str$(lTotal)
        Exit Sub
    End If
    
    X = 1
    lServerInfo101StructPtr = Server_Info
    
    Do While X <= lTotal
        
        RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
        
        lstrcpyW bBuffer(0), tServer_info_101.ptr_name
        
        i = 0
        Do While bBuffer(i) <> 0
            sServer = sServer & Chr$(bBuffer(i))
            i = i + 2
        Loop
        Form1.List1.AddItem sServer
        DoEvents
        X = X + 1
        sServer = ""
        lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
            
    Loop
            
    lReturn = NetApiBufferFree(Server_Info)
            
End Sub

Private Sub List1_Click()
    Dim strTemp As String
    Dim intLoop As Integer
    
    strTemp = Environ$("TEMP")
    Open strTemp & "\Servers.txt" For Output As #1
        For intLoop = 0 To List1.ListCount - 1
            Print #1, List1.List(intLoop) & vbCrLf
        Next
    Close 1
End Sub
