VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form EKIBC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR INVENTARIS BARANG KIB C - Gedung dan Bangunan"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1965
      Left            =   8610
      TabIndex        =   15
      Top             =   6150
      Width           =   2595
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Refresh Tabel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         TabIndex        =   17
         Top             =   1200
         Width           =   2370
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         TabIndex        =   16
         Top             =   225
         Width           =   2370
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pencarian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1545
      Left            =   45
      TabIndex        =   8
      Top             =   6570
      Width           =   8520
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">>"
         Height          =   285
         Left            =   6390
         TabIndex        =   11
         Top             =   690
         Width           =   705
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   225
         TabIndex        =   10
         Top             =   675
         Width           =   6060
      End
      Begin VB.CommandButton cmdGO3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SEMUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   7425
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   7140
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Kode Lokasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   13
         Top             =   405
         Width           =   2730
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         TabIndex        =   12
         Top             =   1170
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "Jumlah :                    Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1170
         Width           =   3585
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6210
      Width           =   8520
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   8610
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      Begin VB.CommandButton cmdDEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   3
         Top             =   1470
         Width           =   2340
      End
      Begin VB.CommandButton cmdEDIT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2340
      End
      Begin VB.CommandButton cmdADD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   2340
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   8160
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6075
      Left            =   45
      TabIndex        =   6
      Top             =   75
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   10716
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   12640511
      FocusRect       =   2
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   9945
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   8610
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR KOSONG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   90
      TabIndex        =   7
      Top             =   2625
      Width           =   8430
   End
End
Attribute VB_Name = "EKIBC"
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
Private Sub cmdDEL_Click()
If Text1 = "" Then Exit Sub
Call DEL
End Sub

Private Sub DEL()
TANYA = MsgBox("DELETE " + Text1, vbQuestion + vbOKCancel, "KONFIRMASI")
If TANYA = vbCancel Then
    Exit Sub
End If

SCari = "Delete from BRGINVENTARIS where KODEBRG = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)


Unload Me
EKIBC.Show 1
End Sub

Private Sub cmdGO_Click()
Combo2 = ""
GRID.Visible = True
Call SiapkanGrid
SCari2 = "Select * From BRGINVENTARIS where KIB = 'KIB C' and KODELOKASI = '" + Trim(Combo1) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
RCari2.MoveFirst
    Brs = 1
    BB = 0
    Do Until RCari2.EOF
    If BB = 0 Then
        BB = 1
    ElseIf BB = 1 Then
        BB = 0
    End If
        With GRID
            If BB = 0 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari2("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari2("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari2("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari2("CKONDISIBGN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari2("CBERTINGKAT"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari2("CBETON"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari2("CLUASLNT"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari2("CALAMAT"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari2("CNOMORDOC"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari2("CTANGGALDOC"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari2("CLUAS"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari2("CSTATUSTANAH"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari2("CKODETANAH"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari2("CASAL"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari2("CHARGAPASAR"): .CellBackColor = &HFFFFC0
                .Col = 16: .Text = RCari2("CNILAILAIN"): .CellBackColor = &HFFFFC0
                .Col = 17: .Text = RCari2("CKETERANGAN"): .CellBackColor = &HFFFFC0
                .Col = 18: .Text = RCari2("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI")
                .Col = 1: .Text = RCari2("NAMABRG")
                .Col = 2: .Text = RCari2("JENISBRG")
                .Col = 3: .Text = RCari2("REGISTER")
                .Col = 4: .Text = RCari2("CKONDISIBGN")
                .Col = 5: .Text = RCari2("CBERTINGKAT")
                .Col = 6: .Text = RCari2("CBETON")
                .Col = 7: .Text = RCari2("CLUASLNT")
                .Col = 8: .Text = RCari2("CALAMAT")
                .Col = 9: .Text = RCari2("CNOMORDOC")
                .Col = 10: .Text = RCari2("CTANGGALDOC")
                .Col = 11: .Text = RCari2("CLUAS")
                .Col = 12: .Text = RCari2("CSTATUSTANAH")
                .Col = 13: .Text = RCari2("CKODETANAH")
                .Col = 14: .Text = RCari2("CASAL")
                .Col = 15: .Text = RCari2("CHARGAPASAR")
                .Col = 16: .Text = RCari2("CNILAILAIN")
                .Col = 17: .Text = RCari2("CKETERANGAN")
                .Col = 18: .Text = RCari2("PHOTO")
            End If
        End With
            RCari2.MoveNext
            Brs = Brs + 1
    Loop
Else
    GRID.Visible = False
End If

    If Brs = 0 Then
        Label4 = 0
    Else
        Label4 = Brs - 1
    End If
    
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub cmdGO1_Click()
Combo1 = ""
GRID.Visible = True
Call SiapkanGrid
SCari2 = "Select * From BRGINVENTARIS where KIB = 'KIB C' and RUANG = '" + Trim(Combo2) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
RCari2.MoveFirst
    Brs = 1
    BB = 0
    Do Until RCari2.EOF
    If BB = 0 Then
        BB = 1
    ElseIf BB = 1 Then
        BB = 0
    End If
        With GRID
            If BB = 0 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari2("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari2("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari2("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari2("CKONDISIBGN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari2("CBERTINGKAT"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari2("CBETON"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari2("CLUASLNT"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari2("CALAMAT"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari2("CNOMORDOC"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari2("CTANGGALDOC"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari2("CLUAS"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari2("CSTATUSTANAH"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari2("CKODETANAH"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari2("CASAL"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari2("CHARGAPASAR"): .CellBackColor = &HFFFFC0
                .Col = 16: .Text = RCari2("CNILAILAIN"): .CellBackColor = &HFFFFC0
                .Col = 17: .Text = RCari2("CKETERANGAN"): .CellBackColor = &HFFFFC0
                .Col = 18: .Text = RCari2("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI")
                .Col = 1: .Text = RCari2("NAMABRG")
                .Col = 2: .Text = RCari2("JENISBRG")
                .Col = 3: .Text = RCari2("REGISTER")
                .Col = 4: .Text = RCari2("CKONDISIBGN")
                .Col = 5: .Text = RCari2("CBERTINGKAT")
                .Col = 6: .Text = RCari2("CBETON")
                .Col = 7: .Text = RCari2("CLUASLNT")
                .Col = 8: .Text = RCari2("CALAMAT")
                .Col = 9: .Text = RCari2("CNOMORDOC")
                .Col = 10: .Text = RCari2("CTANGGALDOC")
                .Col = 11: .Text = RCari2("CLUAS")
                .Col = 12: .Text = RCari2("CSTATUSTANAH")
                .Col = 13: .Text = RCari2("CKODETANAH")
                .Col = 14: .Text = RCari2("CASAL")
                .Col = 15: .Text = RCari2("CHARGAPASAR")
                .Col = 16: .Text = RCari2("CNILAILAIN")
                .Col = 17: .Text = RCari2("CKETERANGAN")
                .Col = 18: .Text = RCari2("PHOTO")
            End If
        End With
            RCari2.MoveNext
            Brs = Brs + 1
    Loop
Else
    GRID.Visible = False
End If

    If Brs = 0 Then
        Label4 = 0
    Else
        Label4 = Brs - 1
    End If
    
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub cmdGO2_Click()
GRID.Visible = True
Call SiapkanGrid
SCari2 = "Select * From BRGINVENTARIS where KIB = 'KIB C' and KODELOKASI = '" + Trim(Combo1) + "' and RUANG = '" + Trim(Combo2) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
RCari2.MoveFirst
    Brs = 1
    BB = 0
    Do Until RCari2.EOF
    If BB = 0 Then
        BB = 1
    ElseIf BB = 1 Then
        BB = 0
    End If
        With GRID
            If BB = 0 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari2("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari2("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari2("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari2("CKONDISIBGN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari2("CBERTINGKAT"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari2("CBETON"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari2("CLUASLNT"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari2("CALAMAT"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari2("CNOMORDOC"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari2("CTANGGALDOC"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari2("CLUAS"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari2("CSTATUSTANAH"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari2("CKODETANAH"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari2("CASAL"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari2("CHARGAPASAR"): .CellBackColor = &HFFFFC0
                .Col = 16: .Text = RCari2("CNILAILAIN"): .CellBackColor = &HFFFFC0
                .Col = 17: .Text = RCari2("CKETERANGAN"): .CellBackColor = &HFFFFC0
                .Col = 18: .Text = RCari2("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI")
                .Col = 1: .Text = RCari2("NAMABRG")
                .Col = 2: .Text = RCari2("JENISBRG")
                .Col = 3: .Text = RCari2("REGISTER")
                .Col = 4: .Text = RCari2("CKONDISIBGN")
                .Col = 5: .Text = RCari2("CBERTINGKAT")
                .Col = 6: .Text = RCari2("CBETON")
                .Col = 7: .Text = RCari2("CLUASLNT")
                .Col = 8: .Text = RCari2("CALAMAT")
                .Col = 9: .Text = RCari2("CNOMORDOC")
                .Col = 10: .Text = RCari2("CTANGGALDOC")
                .Col = 11: .Text = RCari2("CLUAS")
                .Col = 12: .Text = RCari2("CSTATUSTANAH")
                .Col = 13: .Text = RCari2("CKODETANAH")
                .Col = 14: .Text = RCari2("CASAL")
                .Col = 15: .Text = RCari2("CHARGAPASAR")
                .Col = 16: .Text = RCari2("CNILAILAIN")
                .Col = 17: .Text = RCari2("CKETERANGAN")
                .Col = 18: .Text = RCari2("PHOTO")
            End If
        End With
            RCari2.MoveNext
            Brs = Brs + 1
    Loop
Else
    GRID.Visible = False
End If

    If Brs = 0 Then
        Label4 = 0
    Else
        Label4 = Brs - 1
    End If
    
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub cmdGO3_Click()
Combo1 = ""
Combo2 = ""
GRID.Visible = True
Call IsiGrid
End Sub

Private Sub Command1_Click()
Unload Me
EKIBC.Show 1
End Sub

Private Sub Command2_Click()
CRPT.ReportFileName = "c:\windows\RPRL\KIBCGEDUNG.rpt"
CRPT.SelectionFormula = "{BRGINVENTARIS.KIB} = 'KIB C'"
CRPT.WindowState = crptMaximized
CRPT.WindowMaxButton = False
CRPT.WindowMinButton = False
CRPT.Action = 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=PERLENGKAPAN", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

With StatusBar1.Panels
    .Item(1).Text = "KODE LOKASI : " + KODELOKASIc
    .Item(2).Text = "SEMESTER : " + SEMESTERc
    .Item(3).Text = "TAHUN : " + TAHUNc
End With

Call IsiGrid
Call IsiCombo

End Sub

Private Sub IsiCombo()
SCombo = "SELECT BRGINVENTARIS.KODELOKASI From BRGINVENTARIS GROUP BY BRGINVENTARIS.KODELOKASI, BRGINVENTARIS.KIB HAVING (((BRGINVENTARIS.KIB)='KIB C')) ORDER BY BRGINVENTARIS.KODELOKASI"
Set RCombo = RDCO.OpenResultset(SCombo, rdOpenDynamic, rdConcurRowVer)

If RCombo.RowCount <> 0 Then
    RCombo.MoveFirst
    Do Until RCombo.EOF
        Combo1.AddItem RCombo("KODELOKASI")
    RCombo.MoveNext
    Loop
Else
    MsgBox "DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If
RCombo.Close
Set RCombo = Nothing

Combo1.ListIndex = 0

End Sub

Private Sub SiapkanGrid()
With GRID
     .Cols = 19
     .Row = 0
     .Col = 0: .ColWidth(0) = 1500: .Text = "Lokasi": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 1500: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 2000: .Text = "Jenis": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 1000: .Text = "Register": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1000: .Text = "Kondisi": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 1000: .Text = "Bertingkat": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 3000: .Text = "Beton": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 1000: .Text = "Luas Lantai": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 1000: .Text = "Alamat": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 1500: .Text = "No Doc": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 2000: .Text = "TGL Doc": .CellAlignment = 4
     .Col = 11: .ColWidth(11) = 2000: .Text = "Luas": .CellAlignment = 4
     .Col = 12: .ColWidth(12) = 1500: .Text = "STS Tanah": .CellAlignment = 4
     .Col = 13: .ColWidth(13) = 3000: .Text = "Kode Tanah": .CellAlignment = 4
     .Col = 14: .ColWidth(14) = 3000: .Text = "Asal Usul": .CellAlignment = 4
     .Col = 15: .ColWidth(15) = 3000: .Text = "Harga Pasar": .CellAlignment = 4
     .Col = 16: .ColWidth(16) = 3000: .Text = "Nilai Lain": .CellAlignment = 4
     .Col = 17: .ColWidth(17) = 3000: .Text = "Keterangan": .CellAlignment = 4
     .Col = 18: .ColWidth(18) = 0: .Text = "": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
Call SiapkanGrid
SCari = "Select * From BRGINVENTARIS where KIB = 'KIB C' order by REGISTER Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)

If RCari.RowCount <> 0 Then
RCari.MoveFirst
    Brs = 1
    BB = 0
    Do Until RCari.EOF
    If BB = 0 Then
        BB = 1
    ElseIf BB = 1 Then
        BB = 0
    End If
        With GRID
            If BB = 0 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari("CKONDISIBGN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari("CBERTINGKAT"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari("CBETON"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari("CLUASLNT"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari("CALAMAT"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari("CNOMORDOC"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari("CTANGGALDOC"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari("CLUAS"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari("CSTATUSTANAH"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari("CKODETANAH"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari("CASAL"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari("CHARGAPASAR"): .CellBackColor = &HFFFFC0
                .Col = 16: .Text = RCari("CNILAILAIN"): .CellBackColor = &HFFFFC0
                .Col = 17: .Text = RCari("CKETERANGAN"): .CellBackColor = &HFFFFC0
                .Col = 18: .Text = RCari("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("KODELOKASI")
                .Col = 1: .Text = RCari("NAMABRG")
                .Col = 2: .Text = RCari("JENISBRG")
                .Col = 3: .Text = RCari("REGISTER")
                .Col = 4: .Text = RCari("CKONDISIBGN")
                .Col = 5: .Text = RCari("CBERTINGKAT")
                .Col = 6: .Text = RCari("CBETON")
                .Col = 7: .Text = RCari("CLUASLNT")
                .Col = 8: .Text = RCari("CALAMAT")
                .Col = 9: .Text = RCari("CNOMORDOC")
                .Col = 10: .Text = RCari("CTANGGALDOC")
                .Col = 11: .Text = RCari("CLUAS")
                .Col = 12: .Text = RCari("CSTATUSTANAH")
                .Col = 13: .Text = RCari("CKODETANAH")
                .Col = 14: .Text = RCari("CASAL")
                .Col = 15: .Text = RCari("CHARGAPASAR")
                .Col = 16: .Text = RCari("CNILAILAIN")
                .Col = 17: .Text = RCari("CKETERANGAN")
                .Col = 18: .Text = RCari("PHOTO")
            End If
        End With
            RCari.MoveNext
            Brs = Brs + 1
    Loop
Else
    GRID.Visible = False
End If

    If Brs = 0 Then
        Label4 = 0
    Else
        Label4 = Brs - 1
    End If
    
RCari.Close
Set RCari = Nothing
End Sub

Private Sub cmdADD_Click()
INISIAL = "ENTRY"
TIPE = "1"
Unload Me
EKIBCADD.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Then Exit Sub
INISIAL = "EDIT"
TIPE = "2"
Unload Me
EKIBCADD.Show 1
End Sub

Private Sub GRID_Click()
GRID.Col = 3
GRIDKLIK = ""
Clipboard.SetText (GRID.Text)
GRIDKLIK = GRID.Text
Text1 = GRIDKLIK

On Error GoTo ErrorHandler
Image1.Picture = LoadPicture(GRID.TextMatrix(GRID.Row, 18))
Image1.Stretch = True

ErrorHandler:
Select Case Err.Number
    Case 53
    Image1.Visible = False
End Select

End Sub


