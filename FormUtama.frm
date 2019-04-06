VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FormUtama 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yang Ngucapin Ultah"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14430
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   14430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmDataAkhir 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">>"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3070
      Width           =   495
   End
   Begin VB.CommandButton cmDataSelanjutnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3070
      Width           =   495
   End
   Begin VB.CommandButton cmDataSebelumnya 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3070
      Width           =   495
   End
   Begin VB.CommandButton cmDataAwal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3070
      Width           =   495
   End
   Begin VB.CommandButton cmView 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&View"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmTentang 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tentang"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBawah 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   3570
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1680
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmKeluar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmHapus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmBaru 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Baru"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AturKontrol()
    Nyambung
    With Adodc1
        .ConnectionString = CN.ConnectionString
        .RecordSource = "Select * from TableUltah order by Nama Asc;"
        Set DataGrid1.DataSource = Adodc1
        .Refresh
    End With
    With DataGrid1
        .AllowUpdate = False
        .Columns.Item(0).Width = 3000
        .Columns.Item(1).Width = 1000
        .Columns.Item(2).Width = 2000
        .Columns.Item(3).Width = 7700
        .Columns.Item(0).Alignment = dbgLeft
        .Columns.Item(1).Alignment = dbgCenter
        .Columns.Item(2).Alignment = dbgCenter
        .Columns.Item(3).Alignment = dbgLeft
        
        Select Case Month(Date)
            Case Is = 1
                Kalimat = "Januari"
            Case Is = 2
                Kalimat = "Februari"
            Case Is = 3
                Kalimat = "Maret"
            Case Is = 4
                Kalimat = "April"
            Case Is = 5
                Kalimat = "Mei"
            Case Is = 6
                Kalimat = "Juni"
            Case Is = 7
                Kalimat = "Juli"
            Case Is = 8
                Kalimat = "Agustus"
            Case Is = 9
                Kalimat = "September"
            Case Is = 10
                Kalimat = "Oktober"
            Case Is = 11
                Kalimat = "November"
            Case Is = 12
                Kalimat = "Desember"
        End Select
                
        
        StatusBawah.SimpleText = "Yang udah ngucapin : " & Adodc1.Recordset.RecordCount & " orang. | Cell System Count : " & Val(Adodc1.Recordset.RecordCount) * Val(Adodc1.Recordset.Fields.Count) & " Cell(s) | Database Type : MyISAM | Master Record Systems : " & (Val(Adodc1.Recordset.RecordCount) * Val(Adodc1.Recordset.Fields.Count)) * (Val(Adodc1.Recordset.RecordCount)) & " mrs"
    End With
End Sub

Private Sub CmBaru_Click()
    FormBaru.Show vbModal, Me
End Sub

Private Sub cmDataAkhir_Click()
    Adodc1.Recordset.MoveLast
End Sub

Private Sub CmDataAwal_Click()
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmDataSebelumnya_Click()
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF = True Then Adodc1.Recordset.MoveLast
End Sub

Private Sub cmDataSelanjutnya_Click()
    Adodc1.Recordset.MoveNext
    If Adodc1.Recordset.EOF = True Then Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmHapus_Click()
    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "Tidak ada data yang akan dihapus!", vbExclamation + vbOKOnly, "Apa yg dihapus?"
    Else
        X = MsgBox("Anda yakin ingin menghapus data ini?" & vbCrLf & _
                    "-----------------------------------" & vbCrLf & _
                    Adodc1.Recordset.Fields(0).Name & " : " & Adodc1.Recordset.Fields(0).Value & vbCrLf & _
                    Adodc1.Recordset.Fields(1).Name & " : " & Adodc1.Recordset.Fields(1).Value & vbCrLf & _
                    Adodc1.Recordset.Fields(2).Name & " : " & Adodc1.Recordset.Fields(2).Value & vbCrLf & _
                    Adodc1.Recordset.Fields(3).Name & " : " & Adodc1.Recordset.Fields(3).Value & vbCrLf & _
                    "-----------------------------------", vbQuestion + vbYesNo, "Tanya?")
        If X = vbYes Then
            With Adodc1
                .Recordset.Delete
                .Refresh
            End With
            AturKontrol
        End If
    End If
End Sub

Private Sub cmKeluar_Click()
    End
End Sub

Private Sub cmRefresh_Click()
    AturKontrol
End Sub

Private Sub cmTentang_Click()
    MsgBox "''Yang Ngucapin Ultah''" & vbCrLf & _
            "Version 1.0" & vbCrLf & _
            "Programmed by Rizky Khapidsyah" & vbCrLf & _
            "Copyright_(2013) by RikySoft Software House Production", vbInformation + vbOKOnly, "Tentang..."
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub cmView_Click()
    FormView.Show vbModal, Me
End Sub
