VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmGAME 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Everlasting Odyssey"
   ClientHeight    =   7155
   ClientLeft      =   4920
   ClientTop       =   3585
   ClientWidth     =   9585
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmGAME.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGAME.frx":0CCA
   ScaleHeight     =   7155
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrCheckEquips 
      Interval        =   50
      Left            =   6960
      Top             =   1320
   End
   Begin VB.Timer tmrStat 
      Interval        =   50
      Left            =   6480
      Top             =   1320
   End
   Begin VB.CommandButton cmdAddSpeed 
      Caption         =   "+"
      Height          =   375
      Left            =   9120
      TabIndex        =   27
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdAddIntelligence 
      Caption         =   "+"
      Height          =   375
      Left            =   9120
      TabIndex        =   26
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdAddDefence 
      Caption         =   "+"
      Height          =   375
      Left            =   9120
      TabIndex        =   25
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdAddStrength 
      Caption         =   "+"
      Height          =   375
      Left            =   9120
      TabIndex        =   24
      Top             =   5160
      Width           =   375
   End
   Begin VB.Timer tmrDeathCheck 
      Interval        =   50
      Left            =   6480
      Top             =   3000
   End
   Begin VB.CommandButton cmdSouth 
      Caption         =   "South"
      Height          =   615
      Left            =   7440
      TabIndex        =   23
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdNorth 
      Caption         =   "North"
      Height          =   615
      Left            =   7440
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Timer tmrTXTSTATS 
      Interval        =   100
      Left            =   6480
      Top             =   2520
   End
   Begin RichTextLib.RichTextBox txtPlayer 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmGAME.frx":E1D0E
   End
   Begin RichTextLib.RichTextBox txtGame 
      Height          =   6255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   1
      TextRTF         =   $"frmGAME.frx":E1DDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSpeed 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtIntelligence 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox txtDefence 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox txtStrength 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtMana 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtHealth 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtLevel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdEast 
      BackColor       =   &H00000000&
      Caption         =   "East"
      Height          =   735
      Left            =   8640
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   735
      Left            =   6480
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtLocation 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Intelligence:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Defence:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mana:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "       Character"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   9600
      X2              =   6360
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   6360
      Y1              =   3000
      Y2              =   7080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "      Navigation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   9600
      X2              =   6360
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   6360
      Y1              =   3000
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   9600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   6360
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmGAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddDefence_Click()


If StatPoints > 0 Then
 StatPoints = StatPoints - 1
  Defence = Defence + 1
 
 Call AddText("=You Have Increased You Defence By 1 Point!=", MAGENTA, 9, False, True, False)
  
 If StatPoints <> 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Points Left To Spend!=", MAGENTA, 9, False, True, False)
 End If
 
 If StatPoints = 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Point Left To Spend!=", MAGENTA, 9, False, True, False)
 End If
 
 MaxHealth = 100 + (Strength * 1) + (Defence * 3) + Speed
  Health = 100 + (Strength * 1) + (Defence * 3) + Speed
   MP = 20 + (Intelligence * 4)
    MaxMP = 20 + (Intelligence * 4)
    
Call SaveGame
 End If

End Sub

Private Sub cmdAddIntelligence_Click()


If StatPoints > 0 Then
 StatPoints = StatPoints - 1
  Intelligence = Intelligence + 1
 
 Call AddText("=You Have Increased You Intelligence By 1 Point!=", MAGENTA, 9, False, True, False)

 If StatPoints <> 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Points Left To Spend!=", MAGENTA, 9, False, True, False)
 End If
 
 If StatPoints = 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Point Left To Spend!=", MAGENTA, 9, False, True, False)
 End If

 MaxHealth = 100 + (Strength * 1) + (Defence * 3) + Speed
  Health = 100 + (Strength * 1) + (Defence * 3) + Speed
   MP = 20 + (Intelligence * 4)
    MaxMP = 20 + (Intelligence * 4)
    
Call SaveGame
 End If

End Sub

Private Sub cmdAddSpeed_Click()


If StatPoints > 0 Then
 StatPoints = StatPoints - 1
  Speed = Speed + 1
 
 Call AddText("=You Have Increased You Speed By 1 Point!=", MAGENTA, 9, False, True, False)

 If StatPoints <> 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Points Left To Spend!=", MAGENTA, 9, False, True, False)
 End If
 
 If StatPoints = 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Point Left To Spend!=", MAGENTA, 9, False, True, False)
 End If

 MaxHealth = 100 + (Strength * 1) + (Defence * 3) + Speed
  Health = 100 + (Strength * 1) + (Defence * 3) + Speed
   MP = 20 + (Intelligence * 4)
    MaxMP = 20 + (Intelligence * 4)

Call SaveGame
 End If
 
End Sub

Private Sub cmdAddStrength_Click()

If StatPoints > 0 Then
 StatPoints = StatPoints - 1
  Strength = Strength + 1
 
 Call AddText("=You Have Increased You Strength By 1 Point!=", MAGENTA, 9, False, True, False)

 If StatPoints <> 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Points Left To Spend!=", MAGENTA, 9, False, True, False)
 End If
 
 If StatPoints = 1 Then
  Call AddText("=You Have " & StatPoints & " Stat Point Left To Spend!=", MAGENTA, 9, False, True, False)
 End If

 MaxHealth = 100 + (Strength * 1) + (Defence * 3) + Speed
  Health = 100 + (Strength * 1) + (Defence * 3) + Speed
   MP = 20 + (Intelligence * 4)
    MaxMP = 20 + (Intelligence * 4)

Call SaveGame
 End If

End Sub

Private Sub cmdEast_Click()

Call CheckEast

End Sub

Private Sub cmdNorth_Click()

Call CheckNorth


End Sub

Private Sub cmdSouth_Click()

Call CheckSouth

End Sub

Private Sub cmdWest_Click()

Call CheckWest

End Sub

Private Sub Form_Load()


'Game Start Menu
Call AddText("/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\", BRIGHTGREEN, 16, False, False, False)
Call AddText("/\/\/\/\/\/\=Everlasting Odyssey=/\/\/\/\/\/", BRIGHTGREEN, 16, False, False, False)
Call AddText("/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\", BRIGHTGREEN, 16, False, False, False)
Call AddText("==========================", BRIGHTBLUE, 16, False, False, False)
Call AddText("-=+=+=-DoD Productions-=+=+=-", BRIGHTBLUE, 16, True, True, False)
Call AddText("==========================", BRIGHTBLUE, 16, False, False, False)
Call AddText("-------------Copyright 2008-------------", BRIGHTRED, 16, True, True, False)
Call AddText("==========================", BRIGHTRED, 16, False, False, False)


End Sub



Private Sub Form_Unload(Cancel As Integer)
 
End

End Sub



Private Sub tmrCheckEquips_Timer()

Dim pWeapon As String
Dim pHelm As String
Dim pArmor As String
Dim pGloves As String
Dim pLeggings As String
Dim pBoots As String
Dim pAccessory As String

pWeapon = PlayerWeapon
pHelm = PlayerHelm
pArmor = PlayerArmor
pGloves = PlayerGloves
pLeggings = PlayerLeggings
pBoots = PlayerBoots
pAccessory = PlayerAccessory


'check inv for equipped weapon
If PlayerWeapon <> "" Then
 If CheckInvWeapon Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerWeapon)
    PlayerWeapon = ""
     Call AddText("Your " & pWeapon & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If
       
       
       
       
'check inv for equipped helm
If PlayerHelm <> "" Then
 If CheckInvHelm Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerHelm)
    PlayerHelm = ""
     Call AddText("Your " & pHelm & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If




'check inv for equipped armor
If PlayerArmor <> "" Then
 If CheckInvArmor Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerArmor)
    PlayerArmor = ""
     Call AddText("Your " & pArmor & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If




'check inv for equipped gloves
If PlayerGloves <> "" Then
 If CheckInvGloves Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerGloves)
    PlayerGloves = ""
     Call AddText("Your " & pGloves & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If




'check inv for equipped leggings
If PlayerLeggings <> "" Then
 If CheckInvLeggings Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerLeggings)
    PlayerLeggings = ""
     Call AddText("Your " & pLeggings & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If
       
       
       
       
'check inv for equipped boots
If PlayerBoots <> "" Then
 If CheckInvBoots Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerBoots)
    PlayerBoots = ""
     Call AddText("Your " & pBoots & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If
       
       
       
       
'check inv for equipped helm
If PlayerAccessory <> "" Then
 If CheckInvAccessory Then
  
  EquippingItem = False
   Call EquipableItemEffectDatabase(PlayerAccessory)
    PlayerAccessory = ""
     Call AddText("Your " & pAccessory & " Has Been Unequiped And Removed From Your Equipment!", RED, 9, False, True, False)
      Call SaveGame
      
      End If
       End If










End Sub

Private Sub tmrDeathCheck_Timer()

'Check for player death
 If Health <= 0 Then
 Call DoneBattle
  Call AddText("=====", RED, 9, False, False, False)
   Call AddText("You Have Died!!", RED, 9, False, False, True)
    Call AddText("=====", RED, 9, False, False, False)
     Call AddText("", RED, 9, False, False, False)
      
      'Subtract The Experience Penalty + Display It
       Call DeathExperiencePenalty
       
     Health = MaxHealth
      Call SetLocation(ReviveLocation)
       Call UpdateCompass
        Location = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "Location")
         FLocation = LCase(Location)
          Call CheckEvents
           Call SaveGame
     End If



End Sub

Private Sub tmrStat_Timer()

 'Disable Stat + If In Battle
 If Compass = False Then
   cmdAddStrength.Enabled = False
    cmdAddDefence.Enabled = False
     cmdAddIntelligence.Enabled = False
      cmdAddSpeed.Enabled = False
    End If

 
  'Check If Player Has Stat Points To Spend
  If StatPoints > 0 Then
   If Compass = True Then
   cmdAddStrength.Enabled = True
    cmdAddDefence.Enabled = True
     cmdAddIntelligence.Enabled = True
      cmdAddSpeed.Enabled = True
       cmdAddStrength.Visible = True
        cmdAddDefence.Visible = True
         cmdAddIntelligence.Visible = True
          cmdAddSpeed.Visible = True
       End If
        End If
        
  If StatPoints <= 0 Then
   cmdAddStrength.Enabled = False
    cmdAddDefence.Enabled = False
     cmdAddIntelligence.Enabled = False
      cmdAddSpeed.Enabled = False
       cmdAddStrength.Visible = False
        cmdAddDefence.Visible = False
         cmdAddIntelligence.Visible = False
          cmdAddSpeed.Visible = False
       End If
 

End Sub

Private Sub tmrTXTSTATS_Timer()


'Update Stats/Location
txtLevel.Text = Level
 txtHealth.Text = Health
  txtMana.Text = MP
   txtStrength.Text = Strength
    txtDefence.Text = Defence
   txtIntelligence.Text = Intelligence
  txtSpeed.Text = Speed
 txtLocation.Text = Location

'Update Compass
 If Compass = "False" Then
  cmdNorth.Enabled = False
   cmdSouth.Enabled = False
    cmdEast.Enabled = False
   cmdWest.Enabled = False
  End If
  
   If Compass = "True" Then
  cmdNorth.Enabled = True
   cmdSouth.Enabled = True
    cmdEast.Enabled = True
   cmdWest.Enabled = True
  End If
  
  'Update Location
  Location = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "Location")
  FLocation = LCase(Location)
 

 
End Sub

Private Sub txtPlayer_KeyPress(KeyAscii As Integer)

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
     Call TextInput
    End If

End Sub
