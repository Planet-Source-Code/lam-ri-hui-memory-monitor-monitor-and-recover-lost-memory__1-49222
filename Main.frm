VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Monitor"
   ClientHeight    =   3480
   ClientLeft      =   2430
   ClientTop       =   2430
   ClientWidth     =   4800
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4800
   Begin VB.CommandButton Command3 
      Caption         =   "Level 3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Click here to begin level 3 optimization"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Level 2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Click here to begin level 2 optimization"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click here to begin level 1 optimization"
      Top             =   2520
      Width           =   1455
   End
   Begin Memory_Monitor.ProgressCntrl OptmB 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code (C) Lam Ri Hui 2003"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Top             =   3120
      Width           =   2340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Optimize :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Recover_Memory 1
End Sub

Private Sub Command2_Click()
Recover_Memory 2
End Sub

Private Sub Command3_Click()
Recover_Memory 3
End Sub

Private Sub Form_Load()

    Call Always_On_Top(Me.hwnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)
Dim lpBuffer As MemoryStatus
    GlobalMemoryStatus lpBuffer

    lblStatus.Caption = ""
    lblStatus.Caption = lblStatus.Caption & "Available Physical Memory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Total Physical Memory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Used Physical Memory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Percentage Physical Memory : " & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#%") & vbNewLine

End Sub

Private Sub Timer1_Timer()

Dim lpBuffer As MemoryStatus
    GlobalMemoryStatus lpBuffer

    lblStatus.Caption = ""
    lblStatus.Caption = lblStatus.Caption & "Available Physical Memory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Total Physical Memory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Used Physical Memory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Percentage Physical Memory : " & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#%") & vbNewLine

End Sub
Sub Recover_Memory(Level As Integer)
On Error Resume Next
    lblStatus.Caption = "Memory is being optimize. No memory status available."
    Timer1.Enabled = False
    'create an array with 101 value in it
    ReDim a(100) As String
    Dim j As Integer
    OptmB.Max = 100

    For j = 0 To 100
        OptmB.Value = j
        If Level = 1 Then
        a(j) = Space$(500000)
        ElseIf Level = 2 Then
        a(j) = Space$(5000000)
        ElseIf Level = 3 Then
        a(j) = Space$(50000000)
        End If
        DoEvents
        OptmB.Caption = "[" & j / 100 * 100 & "%] Optimizing..."
        
    Next j
    
    For j = 0 To 100
    a(j) = vbNull
    Next j
    Timer1.Enabled = True
    OptmB.Caption = "Done."
    OptmB.Value = 0
    
Dim lpBuffer As MemoryStatus
    GlobalMemoryStatus lpBuffer

    lblStatus.Caption = ""
    lblStatus.Caption = lblStatus.Caption & "Available Physical Memory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Total Physical Memory : " & vbTab & vbTab & Format(CDbl(lpBuffer.dwTotalPhys / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Used Physical Memory : " & vbTab & vbTab & Format(CDbl((lpBuffer.dwTotalPhys - lpBuffer.dwAvailPhys) / 1048576), "#.## MB") & vbNewLine
    lblStatus.Caption = lblStatus.Caption & "Percentage Physical Memory : " & vbTab & Format(CDbl(lpBuffer.dwAvailPhys / lpBuffer.dwTotalPhys), "##.#%") & vbNewLine

End Sub
