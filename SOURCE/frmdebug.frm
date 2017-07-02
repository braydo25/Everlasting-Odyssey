VERSION 5.00
Begin VB.Form frmdebug 
   Caption         =   "Debug"
   ClientHeight    =   3450
   ClientLeft      =   180
   ClientTop       =   570
   ClientWidth     =   6615
   Icon            =   "frmdebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInv2 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtInv1 
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.TextBox txtLocation 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtPText 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtMyText 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Line Line7 
      X1              =   3360
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line6 
      X1              =   3360
      X2              =   0
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   0
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line4 
      X1              =   3360
      X2              =   0
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   3360
      X2              =   0
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Location"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "PText:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "MyText:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "CName:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   1080
      Y2              =   10920
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Debug Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmdebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Text2_Change()

End Sub

Private Sub Timer1_Timer()

txtName.Text = CName
txtMyText.Text = MyText
txtPText.Text = PText
txtLocation.Text = Location




End Sub

