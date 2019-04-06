VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FormView 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Data"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin MSAdodcLib.Adodc AdodcMain 
         Height          =   330
         Left            =   120
         Top             =   840
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
      Begin VB.CommandButton cmDataAkhir 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">>"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmDataSelanjutnya 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">"
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmDataSebelumnya 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton CmDataAwal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox textVia 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   1560
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox textNama 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TextWaktu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   390
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox textIsiPesan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   1095
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "FormView.frx":000C
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton cmTutup 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Tutup"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Isi Pesan"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1440
         TabIndex        =   7
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Melalui (Via)"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   45
      End
   End
End
Attribute VB_Name = "FormView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    MasukkanDataKeView
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .BackColor = Me.BackColor
                .Locked = True
            End With
        End If
    Next
End Sub
Sub MasukkanDataKeView()
    With Me
        .textNama.Text = FormUtama.Adodc1.Recordset.Fields(0).Value
        .textVia.Text = FormUtama.Adodc1.Recordset.Fields(1).Value
        .TextWaktu.Text = FormUtama.Adodc1.Recordset.Fields(2).Value
        .textIsiPesan.Text = FormUtama.Adodc1.Recordset.Fields(3).Value
    End With
End Sub

Private Sub cmDataAkhir_Click()
    FormUtama.Adodc1.Recordset.MoveLast
    MasukkanDataKeView
End Sub

Private Sub CmDataAwal_Click()
    FormUtama.Adodc1.Recordset.MoveFirst
    MasukkanDataKeView
End Sub

Private Sub cmDataSebelumnya_Click()
    FormUtama.Adodc1.Recordset.MovePrevious
    If FormUtama.Adodc1.Recordset.BOF Then FormUtama.Adodc1.Recordset.MoveLast
    MasukkanDataKeView
End Sub

Private Sub cmDataSelanjutnya_Click()
    FormUtama.Adodc1.Recordset.MoveNext
    If FormUtama.Adodc1.Recordset.EOF Then FormUtama.Adodc1.Recordset.MoveFirst
    MasukkanDataKeView
End Sub

Private Sub cmTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub
