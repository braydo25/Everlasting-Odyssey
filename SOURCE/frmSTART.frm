VERSION 5.00
Begin VB.Form frmSTART 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Everlasting Odyssey"
   ClientHeight    =   7155
   ClientLeft      =   6450
   ClientTop       =   8265
   ClientWidth     =   9585
   Icon            =   "frmSTART.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSTART.frx":0CCA
   ScaleHeight     =   7155
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblNew 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "frmSTART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()



End Sub



Private Sub Form_Load()


Call SystemFileChecker


End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub



Private Sub lblLoad_Click()
frmSTART.Visible = False
frmLoad.Visible = True
End Sub

Private Sub lblNew_Click()
frmSTART.Visible = False
frmNEW.Visible = True

End Sub

