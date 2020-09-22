VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption / Decryption"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4320
   Icon            =   "Encryption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "File"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Decrypt"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Encrypt"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Text or data to encrypt / decrypt:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Protection code:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Picture         =   "Encryption.frx":030A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   240
      Picture         =   "Encryption.frx":11D4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PasswordWrong = "&H0"
Private Const Error = "&H1"
Private Sub Command1_Click()
If Text1 = "" Then Text1.SetFocus: Exit Sub
If Text2 = "" Then Text2.SetFocus: Exit Sub

Text2 = Enc.Encrypt(Text2, Text1)
End Sub

Private Sub Command2_Click()
If Text1 = "" Then Text1.SetFocus: Exit Sub
If Text2 = "" Then Text2.SetFocus: Exit Sub

Dim Tempstring As String
Tempstring = Enc.Decrypt(Text2, Text1)
If Tempstring = "ERR1" Then MsgBox "Cannot find encrypted data or data is currupt", vbExclamation, "Error": Exit Sub
If Tempstring = "ERR2" Then MsgBox "Protection password incorrect", vbExclamation, "Error": Text1.SelStart = 0: Text1.SelLength = Len(Text1): Text1.SetFocus: Exit Sub

Text2 = Tempstring
End Sub

Private Sub Command3_Click()
PopupMenu MenuForm.Files
End Sub

Public Sub LoadFile()
Dim Tempstring As String
Dim BinString As String

Tempstring = FileOpen.GetFilename


If Tempstring = "Cancel" Then Exit Sub
Open Tempstring For Binary As #1
BinString = String(FileLen(Tempstring), " ")
Get #1, , BinString
Close #1


Text2 = BinString

End Sub

Public Sub SaveFile()
Dim Tempstring As String

Tempstring = FileOpen.GetFilename

If Tempstring = "cancel" Then Exit Sub

Open Tempstring For Binary As #1
Put #1, , Text2.Text
Close #1



End Sub
