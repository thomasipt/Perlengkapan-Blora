VERSION 5.00
Begin VB.Form LOGON 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "LOGIN "
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4357
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   3443
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4357
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2618
      Width           =   3795
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7492
      TabIndex        =   3
      Top             =   4478
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5677
      TabIndex        =   2
      Top             =   4478
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4357
      TabIndex        =   7
      Top             =   3113
      Width           =   3795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID USER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4357
      TabIndex        =   6
      Top             =   2273
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   5865
      Left            =   52
      Picture         =   "LOGON.frx":0000
      Stretch         =   -1  'True
      Top             =   68
      Width           =   8880
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   974
      TabIndex        =   5
      Top             =   1635
      Width           =   2040
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "USER CODE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   974
      TabIndex        =   4
      Top             =   1245
      Width           =   2040
   End
End
Attribute VB_Name = "LOGON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Public CONEC As String
Private RST As rdoResultset
Private RCari As rdoResultset
Private RSQL As rdoResultset
Private sSQL, SCari As String
Private Sts As String
Private Lolos

Private Sub Label5_Click()
Lolos = 0

sSQL = "Select * From PASSWORD where UserCode = '" + Trim(Text1) + "'and password = '" + Trim(Text2) + "'"
Set RST = RDCO.OpenResultset(sSQL, rdOpenKeyset, rdConcurRowVer)
If RST.RowCount <> 0 Then
    If RST("Status") = 2 Then
        MsgBox "ANDA HARUS MENGGANTI PASSWORD ", vbCritical, "GANTI PASSWORD"
        User = Text1
'        Me.Hide
'        E001.Show
    Exit Sub
    End If
    
  If RST("Status") = 1 Then
        MsgBox "USER ANDA NON AKTIF HUBUNGI ADMINISTRATOR", vbCritical, "NON AKTIF"
        Text2 = ""
        Text2.SetFocus
    Exit Sub
  Else
        CCab = RST("CodeCab")
        Status = RST("Main")
        Operator = RST("UserCode")
        Call Kosong
       
    If Lolos = 1 Then Exit Sub
        Me.Hide
        Unload Me
  End If
Else
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
    MsgBox "AKSES DITOLAK!           ", vbCritical, "PASSWORD"
Exit Sub
End If
RST.Close
Set RST = Nothing

Call KoneksiPERLENGKAPAN

End Sub

Private Sub Label6_Click()
Dim i
i = MsgBox("ANDA YAKIN AKAN KELUAR DARI APLIKASI PERLENGKAPAN ?", vbQuestion + vbOKCancel, "SYSPRL Ver 1.0")
If i = vbOK Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

Call Kosong
Label4 = Date
End Sub

Private Sub Kosong()
Text1 = ""
Text2 = "********"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label6.FontBold = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = True
Label6.FontBold = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label6.FontBold = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label6.FontBold = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label6.FontBold = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label6.FontBold = False
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.FontBold = False
Label6.FontBold = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_GotFocus()
Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Label5_Click
    End If
End Sub


