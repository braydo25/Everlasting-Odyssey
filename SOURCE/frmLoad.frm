VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Everlasting Odyssey"
   ClientHeight    =   3090
   ClientLeft      =   5325
   ClientTop       =   5010
   ClientWidth     =   4680
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCharName 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Character:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter your character's name into the text box and then press load (This Is For Character Security Purposes)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

frmLoad.Visible = False
frmSTART.Visible = True

End Sub


Private Sub cmdLoad_Click()
Dim char As String

char = txtCharName.Text

If Not FileExists("\Characters\" & char & ".EoC") Then
  MsgBox ("That Character Does Not Exist!")
 Exit Sub
End If

If Not FileExists("\Characters\" & char & "Events" & ".EoC") Then
  MsgBox ("That Character Does Not Have An Event Table!")
 Exit Sub
End If

'Character load---------
CName = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Name")
Level = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Level")
Experience = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "Experience")
MaxHealth = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "MaxHealth")
Health = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Health")
MaxMP = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "MaxMana")
MP = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Mana")
Strength = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Strength")
Defence = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Defence")
Intelligence = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Intelligence")
Speed = GetVar(App.Path & "\Characters\" & char & ".Eoc", "Character", "Speed")
Location = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "Location")
Compass = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "Compass")
Gold = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "Gold")
ReviveLocation = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "ReviveLocation")
StatPoints = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "StatPoints")

PlayerHelm = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "PlayerHelm")
PlayerArmor = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "PlayerArmor")
PlayerGloves = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "PlayerGloves")
PlayerLeggings = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "PlayerLeggings")
PlayerBoots = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "PlayerBoots")
PlayerAccessory = GetVar(App.Path & "\Characters\" & char & ".EoC", "Character", "PlayerAccessory")


'----------------------
















FLocation = LCase(Location)

Call LoadInventory

'Event Load------------
BalsoStartRoom = GetVar(App.Path & "\Characters\" & char & "Events" & ".EoC", "Events", "BalsoStartRoom")
BalsoRefugeeFight = GetVar(App.Path & "\Characters\" & char & "Events" & ".EoC", "Events", "BalsoRefugeeFight")
BalsoPortAlura = GetVar(App.Path & "\Characters\" & char & "Events" & ".EoC", "Events", "BalsoPortAlura")
PeasantJacobQuest = GetVar(App.Path & "\Characters\" & char & "Events" & ".EoC", "Events", "PeasantJacobQuest")

If EventExist("VillageOfGlohovaKingAssistant") = True Then
VillageOfGlohovaKingAssistant = GetVar(App.Path & "\Characters\" & char & "Events" & ".EoC", "Events", "VillageOfGlohovaKingAssistant")
End If




If CName <> "" Then
MsgBox ("Load Succssesful!")

ElseIf CName = "" Then
MsgBox ("Error Loading Character!")
Exit Sub
End If


frmLoad.Visible = False
frmGAME.Visible = True

Call CheckEvents

End Sub



Private Sub Form_Unload(Cancel As Integer)

End

End Sub
