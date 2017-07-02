Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub SetCompass(North As String, South As String, West As String, East As String)

frmGAME.cmdNorth.Caption = North
frmGAME.cmdSouth.Caption = South
frmGAME.cmdWest.Caption = West
frmGAME.cmdEast.Caption = East

End Sub

Public Sub SetLocation(Location As String)

Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Location", Location)

End Sub

Public Function GetLocation()

Call GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "Location")

End Function


Public Sub CheckNorth()

'---Balso Village To Balso Throne Room---
If FLocation = "village of balso" Then
Call NullPText
 Call SetLocation("Balso Throne Room")
  Location = "Balso Throne Room"
FLocation = LCase(Location)
   Call AddText("You Are Now In The Balso Throne Room", BROWN, 9, False, True, True)
  Call BalsoThroneRoom
 End If
 
'---Balso Forest To Balso Refugee Camp---

If FLocation = "balso forest" Then
Call NullPText
 Call SetLocation("Balso Refugee Camp")
  Location = "Balso Refugee Camp"
FLocation = LCase(Location)
    Call AddText("You Are Now In The Balso Refugee Camp", BROWN, 9, False, True, True)
   Call BalsoRefugeeCamp
  End If

'---Deep Balso Forest To Balso Forest---

If FLocation = "deep balso forest" Then
Call NullPText
 Call SetLocation("Balso Forest")
  Location = "Balso Forest"
FLocation = LCase(Location)
   Call AddText("You Are Now In The Balso Forest", BROWN, 9, False, True, True)
  Call BalsoForest
 End If

'---Slohovan Shore To Deep Balso Forest---

If FLocation = "slohovan shore" Then
Call NullPText
 Call SetLocation("Deep Balso Forest")
  Location = "Deep Balso Forest"
FLocation = LCase(Location)
 Call AddText("You Are Now In The Deep Balso Forest", BROWN, 9, False, True, True)
Call DeepBalsoForest
End If

'---Port Alura To Slohovan Shore---

If FLocation = "port alura" Then
Call NullPText
 Call SetLocation("Slohovan Shore")
  Location = "Slohovan Shore"
FLocation = LCase(Location)
 Call AddText("You Are Now On The Slohovan Shore", BROWN, 9, False, True, True)
Call SlohovanShore
End If

'---Vile Swamp To Port Glohova---

If FLocation = "vile swamp" Then
Call NullPText
 Call SetLocation("Port Glohova")
  Location = "Port Glohova"
FLocation = LCase(Location)
 Call AddText("You Are Now In Port Glohova", BROWN, 9, False, True, True)
Call PortGlohova
End If

'---Port Glohova To Village Of Glohova---

If FLocation = "port glohova" Then
Call NullPText
 Call SetLocation("Village Of Glohova")
  Location = "Village Of Glohova"
FLocation = LCase(Location)
 Call AddText("You Are Now In The Village Of Glohova", BROWN, 9, False, True, True)
Call VillageOfGlohova
End If

'---Vile Lake To Vile Swamp---

If FLocation = "vile lake" Then
Call NullPText
 Call SetLocation("Vile Swamp")
  Location = "Vile Swamp"
FLocation = LCase(Location)
 Call AddText("You Are Now In The Vile Swamp", BROWN, 9, False, True, True)
Call VileSwamp
End If

End Sub

Public Sub CheckSouth()

'---Balso Throne Room To Village Of Balso---
If FLocation = "balso throne room" Then
Call NullPText
 Location = "Village Of Balso"
FLocation = LCase(Location)
  Call SetLocation("Village Of Balso")
   Call AddText("You Are Now In The Village Of Balso", BROWN, 9, False, True, True)
  Call VillageOfBalso
 End If
 
'---Balso Refugee Camp To Balso Forest---
If FLocation = "balso refugee camp" Then
Call NullPText
 Location = "Balso Forest"
FLocation = LCase(Location)
 Call SetLocation("Balso Forest")
  Call AddText("You Are Now In The Balso Forest", BROWN, 9, False, True, True)
 Call BalsoForest
End If

'---Balso Forest To Deep Balso Forest---
If FLocation = "balso forest" Then
Call NullPText
 Location = "Deep Balso Forest"
FLocation = LCase(Location)
 Call SetLocation("Deep Balso Forest")
  Call AddText("You Are Now In The Deep Balso Forest", BROWN, 9, False, True, True)
 Call DeepBalsoForest
End If

'---Deep Balso Forest To Slohovan Shore---
If FLocation = "deep balso forest" Then
Call NullPText
 Location = "Slohovan Shore"
FLocation = LCase(Location)
 Call SetLocation("Slohovan Shore")
  Call AddText("You Are Now On The Slohovan Shore", BROWN, 9, False, True, True)
 Call SlohovanShore
End If

'---Slohovan Shore To Port Alura---
If FLocation = "slohovan shore" Then
Call NullPText
 Location = "Port Alura"
FLocation = LCase(Location)
 Call SetLocation("Port Alura")
  Call AddText("You Are Now At Port Alura", BROWN, 9, False, True, True)
 Call PortAlura
End If

'---Port Glohova To Vile Swamp---
 If FLocation = "port glohova" Then
 Call NullPText
  Location = "Vile Swamp"
FLocation = LCase(Location)
 Call SetLocation("Vile Swamp")
  Call AddText("You Are Now In The Vile Swamp", BROWN, 9, False, True, True)
 Call VileSwamp
End If
  
'---Village Of Glohova To Port Glohova---
 If FLocation = "village of glohova" Then
 Call NullPText
  Location = "Port Glohova"
FLocation = LCase(Location)
 Call SetLocation("Port Glohova")
  Call AddText("You Are Now In Port Glohova", BROWN, 9, False, True, True)
 Call PortGlohova
End If

'---Vile Swamp To Vile Lake---
 If FLocation = "vile swamp" Then
 Call NullPText
  Location = "Vile Lake"
FLocation = LCase(Location)
 Call SetLocation("Vile Lake")
  Call AddText("You Are Now At The Vile Lake", BROWN, 9, False, True, True)
 Call VileLake
End If
 
 
 
End Sub

Public Sub CheckWest()

'---Balso Forest To Village Of Balso---
If FLocation = "balso forest" Then
Call NullPText
 Location = "Village Of Balso"
FLocation = LCase(Location)
  Call SetLocation("Village Of Balso")
   Call AddText("You Are Now In The Village Of Balso", BROWN, 9, False, True, True)
  Call VillageOfBalso
 End If

'---Slohovan Cape To Slohovan Shore
If FLocation = "slohovan cape" Then
Call NullPText
 Location = "Slohovan Shore"
FLocation = LCase(Location)
  Call SetLocation("Slohovan Shore")
   Call AddText("You Are Now On The Slohovan Shore", BROWN, 9, False, True, True)
  Call SlohovanShore
End If

'---Glohova Cemetary To Village Of Glohova
If FLocation = "glohova cemetary" Then
Call NullPText
 Location = "Village Of Glohova"
FLocation = LCase(Location)
  Call SetLocation("Village Of Glohova")
   Call AddText("You Are Now In The Village Of Glohova", BROWN, 9, False, True, True)
  Call VillageOfGlohova
End If
   
End Sub

Public Sub CheckEast()

'---Village Of Balso To Balso Forest-----
If FLocation = "village of balso" Then
Call NullPText
 Location = "Balso Forest"
FLocation = LCase(Location)
  Call SetLocation("Balso Forest")
   Call AddText("You Are Now In The Balso Forest", BROWN, 9, False, True, True)
  Call BalsoForest
 End If

'---Slohovan Shore to Slohovan Cape---
If FLocation = "slohovan shore" Then
Call NullPText
 Location = "Slohovan Cape"
FLocation = LCase(Location)
  Call SetLocation("Slohovan Cape")
   Call AddText("You Are Now On The Slohovan Cape", BROWN, 9, False, True, True)
  Call SlohovanCape
 End If

'---Village Of Glohova To Glohova Cemetary---
If FLocation = "village of glohova" Then
Call NullPText
 Location = "Glohova Cemetary"
FLocation = LCase(Location)
 Call SetLocation("Glohova Cemetary")
  Call AddText("You Are Now In The Glohova Cemetary", BROWN, 9, False, True, True)
 Call GlohovaCemetary
End If

End Sub

Public Sub UpdateCompass()

 If Compass = "False" Then
  frmGAME.cmdNorth.Enabled = False
   frmGAME.cmdSouth.Enabled = False
    frmGAME.cmdEast.Enabled = False
   frmGAME.cmdWest.Enabled = False
  End If
  
   If Compass = "True" Then
  frmGAME.cmdNorth.Enabled = True
   frmGAME.cmdSouth.Enabled = True
    frmGAME.cmdEast.Enabled = True
   frmGAME.cmdWest.Enabled = True
  End If

End Sub

Public Sub SetReviveLocation(RevLocation As String, RevCost As Integer)

Dim a As Boolean
a = False

 If ReviveLocation <> RevLocation Then
  Call NullPText
  
   Call AddText("Would You Like To Make This Shrine Your New Ressurection Shrine?", MAGENTA, 9, False, True, False)
    Call AddText("The Gods Require " & RevCost & " Gold As A Tribute To Make This Your New Ressurection Shrine.", MAGENTA, 9, False, True, False)
     Call AddText("Would You Like To Tribute " & RevCost & " Gold To The Gods? (Y/N)", MAGENTA, 9, False, True, False)
     
    Do While a = False
    
  If PText = "y" Then
   If Gold >= RevCost Then
    
   Gold = Gold - RevCost
    Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "ReviveLocation", RevLocation)
     ReviveLocation = RevLocation
      Call AddText("", MAGENTA, 9, False, False, False)
       Call AddText("Your Revive Location Has Been Set To " & RevLocation & "!", MAGENTA, 9, False, True, True)
        Call AddText("You Now Have " & Gold & " Gold Left!", MAGENTA, 9, False, True, True)
         Call AddText("", MAGENTA, 9, False, False, False)
          Call SaveGame
           Exit Sub
       
       End If
        End If
        
        
  
    If PText = "y" Then
     If Gold < RevCost Then
      Call AddText("You Do Not Have Enough Gold To Tribute!", MAGENTA, 9, False, True, False)
       Exit Sub
        End If
         End If
  
    
    If PText = "n" Then
     Call AddText("You Chose Not To Set This Shrine As Your Ressurection Shrine!", MAGENTA, 9, False, True, True)
      Exit Sub
    End If
  
  
  
   DoEvents
    Loop
  
  End If
  
  
  
  
 If ReviveLocation = RevLocation Then
  Call NullPText
   Call AddText("", MAGENTA, 9, False, False, False)
    Call AddText("This Is Already Your Current Ressurection Shrine!", MAGENTA, 9, False, True, True)
     Call AddText("", MAGENTA, 9, False, False, False)
      End If
  
End Sub
