VERSION 5.00
Begin VB.Form UPLOAD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPLOAD"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   143
      TabIndex        =   2
      Text            =   "2"
      Top             =   630
      Width           =   4425
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   143
      TabIndex        =   1
      Text            =   "1"
      Top             =   135
      Width           =   4425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   233
      TabIndex        =   0
      Top             =   1125
      Width           =   4245
   End
End
Attribute VB_Name = "UPLOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim a, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String
Private TTL

Private Sub Command1_Click()
IPT = 0
SCari = "Select * from BRGINVENTARIS"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    RCari.MoveFirst
    Do Until RCari.EOF
    KODE_BRG = Format(RCari("KODEBRG"), ">")
        RCari.Edit
            RCari("KODEBRG") = KODE_BRG + "." + IPT
        RCari.Update
        IPT = IPT + 1
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing

MsgBox "SELESAI", vbCritical, "KONFIRMASI"

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

End Sub
