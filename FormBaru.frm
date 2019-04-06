VERSION 5.00
Begin VB.Form FormBaru 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baru.."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Agency FB"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBaru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbVia 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   600
      Width           =   2535
   End
   Begin VB.Timer TimerWaktu 
      Interval        =   10
      Left            =   0
      Top             =   2640
   End
   Begin VB.CommandButton cmBatal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox textKeterangan 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "FormBaru.frx":000C
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox TextWaktu 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox textNama 
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   12
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Melalui (Via)"
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Isi Pesan"
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "FormBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AturKontrol()
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    TextWaktu.Alignment = vbCenter
    With cmbVia
        .Clear
        .AddItem "E-Mail", 0
        .AddItem "SMS", 1
        .AddItem "Facebook", 2
        .AddItem "Twitter", 3
        .AddItem "Alat Komunikasi Lainnya", 4
        .ListIndex = 0
    End With
End Sub
Sub KosongkanInput()
        For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            With Objek
                .Text = ""
                .MaxLength = 254
            End With
        End If
    Next
    textNama.SetFocus
End Sub

Private Sub cmBatal_Click()
    Unload Me
End Sub

Private Sub cmOK_Click()
    If textNama.Text = "" Then
        MsgBox "Silahkan isi nama yang ngucapin ultah kamu!", vbExclamation + vbOKOnly, ""
        textNama.SetFocus
    ElseIf TextWaktu.Text = "" Then
        MsgBox "silahkan isi waktu saat dia ngucapin ultah ke kamu!", vbExclamation + vbOKOnly, ""
        TextWaktu.SetFocus
    ElseIf textKeterangan.Text = "" Then
        MsgBox "Silahkan isi pesan nya!", vbExclamation + vbOKOnly, ""
        textKeterangan.SetFocus
    Else
        With FormUtama.Adodc1
            .Recordset.AddNew
            .Recordset.Fields(0).Value = textNama.Text
            .Recordset.Fields(1).Value = cmbVia.Text
            .Recordset.Fields(2).Value = TextWaktu.Text
            .Recordset.Fields(3).Value = textKeterangan.Text
            .Recordset.Update
            .Refresh
        End With
        KosongkanInput
        FormUtama.AturKontrol
        cmBatal.Caption = "&Tutup"
    End If
End Sub

Private Sub Form_Load()
    AturKontrol
End Sub

Private Sub TimerWaktu_Timer()
    TextWaktu.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date) & " - " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
End Sub
