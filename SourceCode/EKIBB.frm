VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form EKIBB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR INVENTARIS BARANG KIB B - Peralatan dan Mesin"
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
      TabIndex        =   14
      Top             =   6150
      Width           =   2595
      Begin VB.CommandButton Command2 
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
      Begin VB.CommandButton Command1 
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
         TabIndex        =   15
         Top             =   1200
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
      TabIndex        =   7
      Top             =   6570
      Width           =   8520
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   225
         TabIndex        =   17
         Top             =   675
         Width           =   6195
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
      Begin VB.CommandButton cmdGO 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">>"
         Height          =   285
         Left            =   6525
         TabIndex        =   8
         Top             =   690
         Width           =   705
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
         Left            =   2205
         TabIndex        =   11
         Top             =   405
         Width           =   2730
      End
      Begin VB.Shape Shape1 
         Height          =   750
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   7275
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
         TabIndex        =   10
         Top             =   1170
         Width           =   3585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   8610
      TabIndex        =   5
      Top             =   0
      Width           =   2595
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
         TabIndex        =   0
         Top             =   210
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
         TabIndex        =   1
         Top             =   840
         Width           =   2340
      End
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
         TabIndex        =   2
         Top             =   1470
         Width           =   2340
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6210
      Width           =   8520
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
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
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6588
            TextSave        =   ""
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
      TabIndex        =   13
      Top             =   90
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
      Left            =   9360
      Top             =   2655
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
      Left            =   45
      TabIndex        =   6
      Top             =   2632
      Width           =   8430
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   8610
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2595
   End
End
Attribute VB_Name = "EKIBB"
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
EKIBB.Show 1
End Sub

Private Sub cmdGO_Click()
Combo2 = ""
GRID.Visible = True
Call SiapkanGrid
SCari2 = "Select * From BRGINVENTARIS where KIB = 'KIB B' and KODELOKASI = '" + Trim(Combo1) + "'"
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
                .Col = 0: .Text = RCari("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari2("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari2("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari2("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari2("TAHUN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari2("BMERK"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari2("BUKURAN"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari2("BBAHAN"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari2("BPABRIK"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari2("BRANGKA"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari2("BMESIN"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari2("BPOLISI"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari2("BBPKB"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari2("BASAL"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari2("BHARGA"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari2("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI")
                .Col = 1: .Text = RCari2("REGISTER")
                .Col = 2: .Text = RCari2("NAMABRG")
                .Col = 3: .Text = RCari2("JENISBRG")
                .Col = 4: .Text = RCari2("TAHUN")
                .Col = 5: .Text = RCari2("BMERK")
                .Col = 6: .Text = RCari2("BUKURAN")
                .Col = 7: .Text = RCari2("BBAHAN")
                .Col = 8: .Text = RCari2("BPABRIK")
                .Col = 9: .Text = RCari2("BRANGKA")
                .Col = 10: .Text = RCari2("BMESIN")
                .Col = 11: .Text = RCari2("BPOLISI")
                .Col = 12: .Text = RCari2("BBPKB")
                .Col = 13: .Text = RCari2("BASAL")
                .Col = 14: .Text = RCari2("BHARGA")
                .Col = 15: .Text = RCari2("PHOTO")
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
SCari2 = "Select * From BRGINVENTARIS where KIB = 'KIB B' and RUANG = '" + Trim(Combo2) + "'"
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
                .Col = 0: .Text = RCari2("KODELOKASI"): .CellBackColor = &HFFFFC0
                .Col = 1: .Text = RCari2("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari2("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari2("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari2("TAHUN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari2("BMERK"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari2("BUKURAN"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari2("BBAHAN"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari2("BPABRIK"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari2("BRANGKA"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari2("BMESIN"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari2("BPOLISI"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari2("BBPKB"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari2("BASAL"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari2("BHARGA"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari2("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI")
                .Col = 1: .Text = RCari2("REGISTER")
                .Col = 2: .Text = RCari2("NAMABRG")
                .Col = 3: .Text = RCari2("JENISBRG")
                .Col = 4: .Text = RCari2("TAHUN")
                .Col = 5: .Text = RCari2("BMERK")
                .Col = 6: .Text = RCari2("BUKURAN")
                .Col = 7: .Text = RCari2("BBAHAN")
                .Col = 8: .Text = RCari2("BPABRIK")
                .Col = 9: .Text = RCari2("BRANGKA")
                .Col = 10: .Text = RCari2("BMESIN")
                .Col = 11: .Text = RCari2("BPOLISI")
                .Col = 12: .Text = RCari2("BBPKB")
                .Col = 13: .Text = RCari2("BASAL")
                .Col = 14: .Text = RCari2("BHARGA")
                .Col = 15: .Text = RCari2("PHOTO")
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
SCari2 = "Select * From BRGINVENTARIS where KIB = 'KIB B' and KODELOKASI = '" + Trim(Combo1) + "' and RUANG = '" + Trim(Combo2) + "'"
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
                .Col = 1: .Text = RCari2("REGISTER"): .CellBackColor = &HFFFFC0
                .Col = 2: .Text = RCari2("NAMABRG"): .CellBackColor = &HFFFFC0
                .Col = 3: .Text = RCari2("JENISBRG"): .CellBackColor = &HFFFFC0
                .Col = 4: .Text = RCari2("TAHUN"): .CellBackColor = &HFFFFC0
                .Col = 5: .Text = RCari2("BMERK"): .CellBackColor = &HFFFFC0
                .Col = 6: .Text = RCari2("BUKURAN"): .CellBackColor = &HFFFFC0
                .Col = 7: .Text = RCari2("BBAHAN"): .CellBackColor = &HFFFFC0
                .Col = 8: .Text = RCari2("BPABRIK"): .CellBackColor = &HFFFFC0
                .Col = 9: .Text = RCari2("BRANGKA"): .CellBackColor = &HFFFFC0
                .Col = 10: .Text = RCari2("BMESIN"): .CellBackColor = &HFFFFC0
                .Col = 11: .Text = RCari2("BPOLISI"): .CellBackColor = &HFFFFC0
                .Col = 12: .Text = RCari2("BBPKB"): .CellBackColor = &HFFFFC0
                .Col = 13: .Text = RCari2("BASAL"): .CellBackColor = &HFFFFC0
                .Col = 14: .Text = RCari2("BHARGA"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari2("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari2("KODELOKASI")
                .Col = 1: .Text = RCari2("REGISTER")
                .Col = 2: .Text = RCari2("NAMABRG")
                .Col = 3: .Text = RCari2("JENISBRG")
                .Col = 4: .Text = RCari2("TAHUN")
                .Col = 5: .Text = RCari2("BMERK")
                .Col = 6: .Text = RCari2("BUKURAN")
                .Col = 7: .Text = RCari2("BBAHAN")
                .Col = 8: .Text = RCari2("BPABRIK")
                .Col = 9: .Text = RCari2("BRANGKA")
                .Col = 10: .Text = RCari2("BMESIN")
                .Col = 11: .Text = RCari2("BPOLISI")
                .Col = 12: .Text = RCari2("BBPKB")
                .Col = 13: .Text = RCari2("BASAL")
                .Col = 14: .Text = RCari2("BHARGA")
                .Col = 15: .Text = RCari2("PHOTO")
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
EKIBB.Show 1
End Sub

Private Sub Command2_Click()
Crpt.ReportFileName = "c:\windows\RPRL\KIBBPERALATAN.rpt"
Crpt.SelectionFormula = "{BRGINVENTARIS.KIB} = 'KIB B'"
Crpt.WindowState = crptMaximized
Crpt.WindowMaxButton = False
Crpt.WindowMinButton = False
Crpt.Action = 1
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
SCombo = "SELECT BRGINVENTARIS.KODELOKASI From BRGINVENTARIS GROUP BY BRGINVENTARIS.KODELOKASI, BRGINVENTARIS.KIB HAVING (((BRGINVENTARIS.KIB)='KIB B')) ORDER BY BRGINVENTARIS.KODELOKASI"
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

SCombo2 = "Select RUANG from V_RUANG order by RUANG"
Set RCombo2 = RDCO.OpenResultset(SCombo2, rdOpenDynamic, rdConcurRowVer)

Combo1.ListIndex = 0

End Sub

Private Sub SiapkanGrid()
With GRID
     .Cols = 16
     .Row = 0
     .Col = 0: .ColWidth(0) = 1500: .Text = "Lokasi": .CellAlignment = 4
     .Col = 1: .ColWidth(1) = 1500: .Text = "Register": .CellAlignment = 4
     .Col = 2: .ColWidth(2) = 2000: .Text = "Nama Barang": .CellAlignment = 4
     .Col = 3: .ColWidth(3) = 2000: .Text = "JNS Barang": .CellAlignment = 4
     .Col = 4: .ColWidth(4) = 1000: .Text = "Tahun": .CellAlignment = 4
     .Col = 5: .ColWidth(5) = 2000: .Text = "Merk": .CellAlignment = 4
     .Col = 6: .ColWidth(6) = 1000: .Text = "Ukuran": .CellAlignment = 4
     .Col = 7: .ColWidth(7) = 1000: .Text = "Bahan": .CellAlignment = 4
     .Col = 8: .ColWidth(8) = 1000: .Text = "Pabrik": .CellAlignment = 4
     .Col = 9: .ColWidth(9) = 1500: .Text = "Rangka": .CellAlignment = 4
     .Col = 10: .ColWidth(10) = 2000: .Text = "No Mesin": .CellAlignment = 4
     .Col = 11: .ColWidth(11) = 2000: .Text = "No Polisi": .CellAlignment = 4
     .Col = 12: .ColWidth(12) = 1500: .Text = "No BBPKB": .CellAlignment = 4
     .Col = 13: .ColWidth(13) = 3000: .Text = "Asal": .CellAlignment = 4
     .Col = 14: .ColWidth(14) = 1000: .Text = "Harga": .CellAlignment = 4
     .Col = 15: .ColWidth(15) = 0: .Text = "": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
Call SiapkanGrid
SCari = "Select * From BRGINVENTARIS where KIB = 'KIB B' order by REGISTER Asc"
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
                .Col = 0: .Text = RCari("KODELOKASI"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 1: .Text = RCari("REGISTER"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 2: .Text = RCari("NAMABRG"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 3: .Text = RCari("JENISBRG"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 4: .Text = RCari("TAHUN"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 5: .Text = RCari("BMERK"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 6: .Text = RCari("BUKURAN"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 7: .Text = RCari("BBAHAN"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 8: .Text = RCari("BPABRIK"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 9: .Text = RCari("BRANGKA"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 10: .Text = RCari("BMESIN"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 11: .Text = RCari("BPOLISI"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 12: .Text = RCari("BBPKB"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 13: .Text = RCari("BASAL"): .CellBackColor = &HFFFFC0: .CellAlignment = 3
                .Col = 14: .Text = Format(RCari("BHARGA"), "##,###.00"): .CellBackColor = &HFFFFC0
                .Col = 15: .Text = RCari("PHOTO"): .CellBackColor = &HFFFFC0
            ElseIf BB = 1 Then
                .Rows = Brs + 1
                .Row = Brs
                .Col = 0: .Text = RCari("KODELOKASI"): .CellAlignment = 3
                .Col = 1: .Text = RCari("REGISTER"): .CellAlignment = 3
                .Col = 2: .Text = RCari("NAMABRG"): .CellAlignment = 3
                .Col = 3: .Text = RCari("JENISBRG"): .CellAlignment = 3
                .Col = 4: .Text = RCari("TAHUN"): .CellAlignment = 3
                .Col = 5: .Text = RCari("BMERK"): .CellAlignment = 3
                .Col = 6: .Text = RCari("BUKURAN"): .CellAlignment = 3
                .Col = 7: .Text = RCari("BBAHAN"): .CellAlignment = 3
                .Col = 8: .Text = RCari("BPABRIK"): .CellAlignment = 3
                .Col = 9: .Text = RCari("BRANGKA"): .CellAlignment = 3
                .Col = 10: .Text = RCari("BMESIN"): .CellAlignment = 3
                .Col = 11: .Text = RCari("BPOLISI"): .CellAlignment = 3
                .Col = 12: .Text = RCari("BBPKB"): .CellAlignment = 3
                .Col = 13: .Text = RCari("BASAL"): .CellAlignment = 3
                .Col = 14: .Text = Format(RCari("BHARGA"), "##,###.00")
                .Col = 15: .Text = RCari("PHOTO")
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
EKIBBADD.Show 1
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Then Exit Sub
INISIAL = "EDIT"
TIPE = "2"
Unload Me
EKIBBADD.Show 1
End Sub

Private Sub GRID_Click()
GRID.Col = 1
GRIDKLIK = ""
Clipboard.SetText (GRID.Text)
GRIDKLIK = GRID.Text
Text1 = GRIDKLIK

On Error GoTo ErrorHandler
Image1.Picture = LoadPicture(GRID.TextMatrix(GRID.Row, 15))
Image1.Stretch = True

ErrorHandler:
Select Case Err.Number
    Case 53
    Image1.Visible = False
End Select

End Sub


