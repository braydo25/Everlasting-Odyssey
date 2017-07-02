VERSION 5.00
Begin VB.Form frmNEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Everlasting Odyssey"
   ClientHeight    =   5850
   ClientLeft      =   4410
   ClientTop       =   3915
   ClientWidth     =   4935
   Icon            =   "frmNEW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNEW.frx":0CCA
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdResetStats 
      Caption         =   "Reset Stats"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   26
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4440
      Top             =   120
   End
   Begin VB.CommandButton cmdSubSpeed 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdSubIntelligence 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdSubDefence 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   23
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdSubStrength 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdAddSpeed 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdAddIntelligence 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdAddDefence 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdAddStrength 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txtSpendable 
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtSpeed 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtIntelligence 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtDefence 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtStrength 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtMana 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtHealth 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtNewName 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   328
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   328
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   328
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   328
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   328
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   328
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   328
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Spendable Stats:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblIntelligence 
      BackStyle       =   0  'Transparent
      Caption         =   "Intelligence:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblDefence 
      BackStyle       =   0  'Transparent
      Caption         =   "Defence:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblStrength 
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "Mana:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line8 
      X1              =   -8
      X2              =   328
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAddDefence_Click()

If NewDefence >= 15 Then
 Call MsgBox("This Stat Cannot Be Raised Further At This Time!")
  Exit Sub
   End If
   
If NewSpendable <= 0 Then
 Call MsgBox("You Have 0 Spendable Stat Points Left!")
  Exit Sub
   End If
   
If NewDefence < 15 Then
 If NewSpendable > 0 Then
  NewDefence = NewDefence + 1
   NewSpendable = NewSpendable - 1
   End If
    End If

End Sub

Private Sub cmdAddIntelligence_Click()

If NewIntelligence >= 15 Then
 Call MsgBox("This Stat Cannot Be Raised Further At This Time!")
  Exit Sub
   End If
   
If NewSpendable <= 0 Then
 Call MsgBox("You Have 0 Spendable Stat Points Left!")
  Exit Sub
   End If
   
If NewIntelligence < 15 Then
 If NewSpendable > 0 Then
  NewIntelligence = NewIntelligence + 1
   NewSpendable = NewSpendable - 1
   End If
    End If

End Sub

Private Sub cmdAddSpeed_Click()

If NewSpeed >= 15 Then
 Call MsgBox("This Stat Cannot Be Raised Further At This Time!")
  Exit Sub
   End If
   
If NewSpendable <= 0 Then
 Call MsgBox("You Have 0 Spendable Stat Points Left!")
  Exit Sub
   End If
   
If NewSpeed < 15 Then
 If NewSpendable > 0 Then
  NewSpeed = NewSpeed + 1
   NewSpendable = NewSpendable - 1
   End If
    End If

End Sub

Private Sub cmdAddStrength_Click()

If NewStrength >= 15 Then
 Call MsgBox("This Stat Cannot Be Raised Further At This Time!")
  Exit Sub
   End If
   
If NewSpendable <= 0 Then
 Call MsgBox("You Have 0 Spendable Stat Points Left!")
  Exit Sub
   End If
   
If NewStrength < 15 Then
 If NewSpendable > 0 Then
  NewStrength = NewStrength + 1
   NewSpendable = NewSpendable - 1
   End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
frmNEW.Visible = False
frmSTART.Visible = True




End Sub

Private Sub cmdConfirm_Click()
        

  If txtNewName.Text = "" Then
   Call MsgBox("You Have Not Entered A Name!")
    Exit Sub
     End If
      
  If NewSpendable > 0 Then
   Call MsgBox("You Have Not Spent All Your Stat Points!")
    Exit Sub
     End If
      
      CName = txtNewName.Text
    
    
'Make Sure There Isn't A Char Already Named That
If FileExists("\Characters\" & CName & ".EoC") Then
MsgBox ("That Character Name Is In Use!")
Exit Sub
End If
'-----------------------------------------------

Level = 1
Experience = 0
MaxHealth = NewMaxHealth
Health = NewHealth
MaxMP = NewMaxMP
MP = NewMP
Strength = NewStrength
Defence = NewDefence
Intelligence = NewIntelligence
Speed = NewSpeed
Location = "Balso Throne Room"
FLocation = LCase(Location)
Gold = 0
ReviveLocation = "Balso Throne Room"
StatPoints = 0

'Make Char Event Flags
Call WriteEoC("Events", "BalsoStartRoom", "0", (App.Path & "\Characters\" & CName & "Events" & ".EoC"))
Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "BalsoRefugeeFight", "0")
Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "BalsoPortAlura", "0")
Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "PeasantJacobQuest", "0")


'Make Char Data And File
  Call WriteEoC("Character", "Name", CName, (App.Path & "\Characters\" & CName & ".EoC"))
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Level", "1")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Experience", "0")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "MaxHealth", MaxHealth)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Health", Health)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "MaxMana", MaxMP)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Mana", MP)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Strength", Strength)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Defence", Defence)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Intelligence", Intelligence)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Speed", Speed)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Location", "Balso Throne Room")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Compass", "False")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Gold", "0")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "ReviveLocation", "Village Of Balso")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "StatPoints", "0")
  
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerHelm", "")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerArmor", "")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerGloves", "")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerLeggings", "")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerBoots", "")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerAccessory", "")
  
  
  'place inventory
  Call InputInvOnNew
  
MsgBox ("Your Character Has Been Created!")

frmNEW.Visible = False
frmGAME.Visible = True

Call BalsoThroneRoom

End Sub

Private Sub cmdResetStats_Click()

NewSpendable = 30
NewStrength = 0
NewDefence = 0
NewIntelligence = 0
NewSpeed = 0

End Sub

Private Sub cmdSubDefence_Click()

If NewDefence <= 0 Then
 Call MsgBox("There Are 0 Removable Stats In This Category!")
  Exit Sub
   End If
   
If NewDefence > 0 Then
 NewDefence = NewDefence - 1
  NewSpendable = NewSpendable + 1
   Exit Sub
    End If
    

End Sub

Private Sub cmdSubIntelligence_Click()

If NewIntelligence <= 0 Then
 Call MsgBox("There Are 0 Removable Stats In This Category!")
  Exit Sub
   End If
   
If NewIntelligence > 0 Then
 NewIntelligence = NewIntelligence - 1
  NewSpendable = NewSpendable + 1
   Exit Sub
    End If
    

End Sub

Private Sub cmdSubSpeed_Click()

If NewSpeed <= 0 Then
 Call MsgBox("There Are 0 Removable Stats In This Category!")
  Exit Sub
   End If
   
If NewSpeed > 0 Then
 NewSpeed = NewSpeed - 1
  NewSpendable = NewSpendable + 1
   Exit Sub
    End If
    

End Sub

Private Sub cmdSubStrength_Click()

If NewStrength <= 0 Then
 Call MsgBox("There Are 0 Removable Stats In This Category!")
  Exit Sub
   End If
   
If NewStrength > 0 Then
 NewStrength = NewStrength - 1
  NewSpendable = NewSpendable + 1
   Exit Sub
    End If
    

End Sub



Private Sub Form_Load()


NewSpendable = 30

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub Timer1_Timer()

NewMaxHealth = 100 + (NewStrength * 1) + (NewDefence * 3) + NewSpeed
NewMP = 20 + (NewIntelligence * 4)

txtSpendable.Text = NewSpendable
txtHealth.Text = NewMaxHealth
txtMana.Text = NewMP
txtStrength.Text = NewStrength
txtDefence.Text = NewDefence
txtIntelligence.Text = NewIntelligence
txtSpeed.Text = NewSpeed
NewHealth = NewMaxHealth

End Sub
