VERSION 5.00
Begin VB.Form MenuForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MenuForm"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   3900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Files 
      Caption         =   "Files"
      Begin VB.Menu Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Load 
         Caption         =   "&Load"
      End
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Load_Click()
Form1.LoadFile

End Sub

Private Sub Save_Click()
Form1.SaveFile
End Sub
