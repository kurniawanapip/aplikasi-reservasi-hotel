VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtpcekout 
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   96337921
      CurrentDate     =   42797
   End
   Begin MSComCtl2.DTPicker dtpcekin 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   96337921
      CurrentDate     =   42797
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdhitung 
      Caption         =   "Hitung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox cmbkelas 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":000D
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblbayar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lbltarif 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image picfoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah Pembayaran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Cek-Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Cek-In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Kelas Kamar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "PROGRAM APLIKASI RESERVASI HOTEL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vbayar, vtarif, vlamainap As Single

Private Sub cmbkelas_Click()
vPath = "D:\img\" & cmbkelas.Text & ".jpg"
If Dir$(vPath) <> "" Then
      picfoto.Picture = LoadPicture(vPath)
Else
      picfoto.Picture = LoadPicture("")
End If
End Sub

Private Sub cmdhapus_Click()
lblbayar = ""
lbltarif = ""
txtnama = ""
cmbkelas = ""
txtnama.SetFocus
End Sub

Private Sub cmdhitung_Click()
Select Case cmbkelas.Text
Case "de-lux": vtarif = 400000
Case "de-suite": vtarif = 350000
Case "vip": vtarif = 500000
End Select

vlamainap = dtpcekout.Value - dtpcekin.Value
If vlamainap < 2 Then
vbayar = vtarif * 2
Else
vbayar = vtarif * vlamainap
End If

lbltarif.Caption = Format(vtarif, "currency")
lblbayar.Caption = Format(vbayar, "currency")

End Sub

