Attribute VB_Name = "modGAMEDATA"
Option Explicit

Public Sub SaveGame()


'Character Save
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Level", Level)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Experience", Experience)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "MaxHealth", MaxHealth)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Health", Health)
  Call PutVar((App.Path & "\characters\" & CName & ".EoC"), "Character", "MaxMana", MaxMP)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Mana", MP)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Strength", Strength)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Defence", Defence)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Intelligence", Intelligence)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Speed", Speed)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Location", Location)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Gold", Gold)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "ReviveLocation", ReviveLocation)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "StatPoints", StatPoints)
'Set Compass
 
 If Compass = True Then
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Compass", "True")
   End If
   
 If Compass = False Then
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Compass", "False")
   End If
   
'save equipment
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerHelm", PlayerHelm)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerArmor", PlayerArmor)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerGloves", PlayerGloves)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerLeggings", PlayerLeggings)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerBoots", PlayerBoots)
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "PlayerAccessory", PlayerAccessory)
   
   
'save inventory
Call SaveInventory
'---

Call AddText("-Your Game Has Been Saved-", PINK, 8, False, True, True)



End Sub

Public Sub BalsoThroneRoom()



'Chat Sequence---------
If BalsoStartRoom = 0 Then
 Call AddText("King Balso: Hello there young adventurer, it seems you have finally awakened from your long slumber. If you can't remember, my soldiers found you on the Slovohan Cape.", BRIGHTBLUE, 9, False, False, False)
  Call AddText("An Injured Soldier Suddenly Rushes In", BLUE, 9, True, False, False)
   Call AddText("Injured Soldier: Sire! Some bandits are attacking the refugee camp outside of our northern walls! We must send more soldiers!", BRIGHTBLUE, 9, False, False, False)
    Call AddText("The King stops and thinks...", BLUE, 9, True, False, False)
     Call AddText("King Balso: All of our available troops have left to investigate the strange activities recently happening on Mount. Dragona...", BRIGHTBLUE, 9, False, False, False)
      Call AddText("Another Injured Soldier Rushes In..", BLUE, 9, True, False, False)
       Call AddText("Injured Soldier(2): Sire!! We need more soldiers to assist fending off these foes from our refugee camp!", BRIGHTBLUE, 9, False, False, False)
      Call AddText("King Balso: We have no available soldiers! However... " & CName & ", Would you go to the refugee camp and assist our other soldiers in warding off this foe? (Y/N)", BRIGHTBLUE, 9, False, True, False)
    

'wait for input

Do While BalsoStartRoom = 0

 
   
'yes answer
 If PText = "y" Then
  Call AddText("King Balso: Excellent! Then off to the refugee camp with you! It shouldn't be to hard to find..", BRIGHTBLUE, 9, False, False, False)
   Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "BalsoStartRoom", "1")
    Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Compass", "True")
     Compass = True
      Call SetCompass("", "Village Of Balso", "", "")
     BalsoStartRoom = 1
   Exit Sub
  End If
           'no answer
           If PText = "n" Then
            Call AddText("King Balso: No? You tell me, King Balso.. NO!? You had better go assist my soldiers, otherwise I'll have your head on a stick by the end of the morning! GET TO THE CAMP! NOW!", BRIGHTBLUE, 9, False, False, False)
             Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "BalsoStartRoom", "1")
              Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "Compass", "True")
               Compass = True
                Call SetCompass("", "Village Of Balso", "", "")
               BalsoStartRoom = 1
             Exit Sub
            End If
  
  DoEvents
 Loop
End If

'Throne Room Commands------------
Call BalsoKingTalk




End Sub

Public Sub VillageOfBalso()

Dim i As Boolean
Dim RandSpeech As Integer

i = False

Call SetCompass("Castle Of Balso", "", "", "Balso Forest")

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To Villagers  /  Shop / Use Ressurection Shrine", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)

'----Talk To Villagers----



'wait for input
Do While i = False
 
 If FLocation <> "village of balso" Then
  Exit Sub
   End If
 
 'check for talk to villagers
 If PText = "talk to villagers" Then
  
   RandSpeech = Rand(1, 6)
   
  
  '1st villager
  If RandSpeech = 1 Then
   Call AddText("Random Villager: I heard from my father that creatures carry gold! How strange is that?", BRIGHTBLUE, 9, False, False, False)
    Call NullPText

   End If

  '2nd villager
  If RandSpeech = 2 Then
   Call AddText("Random Villager: I wish there was more to do in this village.. maybe the Capital has more to do?", BRIGHTBLUE, 9, False, False, False)
    Call NullPText

   End If
   
  '3rd villager
  If RandSpeech = 3 Then
   Call AddText("Random Villager: I want to be a Knight when I get older!", BRIGHTBLUE, 9, False, False, False)
    Call NullPText

   End If
    
  '4th villager
  If RandSpeech = 4 Then
   Call AddText("Random Villager: My grandfather tells me stories of other worlds, with creatures that towered above the treetops, and a race of people that were no taller than a newborn child!", BRIGHTBLUE, 9, False, False, False)
    Call NullPText

   End If
   
   '5th villager
   If RandSpeech = 5 Then
    Call AddText("Random Villager: It seems that the Capital's Mages find new ways to use magic everyday!", BRIGHTBLUE, 9, False, False, False)
     Call NullPText

    End If
    
   '6th villager
   If RandSpeech = 6 Then
    Call AddText("Random Villager: This village is so boriiiiing...", BRIGHTBLUE, 9, False, False, False)
     Call NullPText

    End If
End If

'set revive location

 If PText = "use ressurection shrine" Then
  Call SetReviveLocation("Village Of Balso", 50)
   End If

 
'shop
If PText = "shop" Then
 Call Shop("Potion", "Soul Potion", "Copper Sword", "Hide Chestplate", "Hide Leggings", "Fur Gloves", "Fur Boots", "", "", "")
  Call NullPText
   End If

 
 
 
 
DoEvents
Loop



End Sub

Public Sub BalsoForest()

Dim a As Boolean
a = False

Call SetCompass("Balso Refugee Camp", "Deep Balso Forest", "Village of Balso", "")
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)


'hunt command
Do While a = False

If FLocation <> "balso forest" Then
 Exit Sub
End If

If PText = "hunt" Then
 Call Hunt("Bandit", "Wolf", "Ferocious Rat", "Treetop Spitter", "Bandit", "Wolf", "Ferocious Rat", "Treetop Spitter", "Bandit")
  Exit Sub
End If

DoEvents
Loop
'--------

End Sub


Public Sub CheckEvents()

'--Balso Throne Room
 If FLocation = "balso throne room" Then
  Call BalsoThroneRoom
   Exit Sub
  End If
'-------------------

'--Village Of Balso
 If FLocation = "village of balso" Then
  Call VillageOfBalso
   Exit Sub
 End If
'-------------------

'--Balso Forest
 If FLocation = "balso forest" Then
  Call BalsoForest
   Exit Sub
 End If
'-------------------

'--Balso Refugee Camp
 If FLocation = "balso refugee camp" Then
  If BalsoRefugeeInFight = False Then
   Call BalsoRefugeeCamp
    Exit Sub
 End If
End If
'-------------------

'--Deep Balso Forest
 If FLocation = "deep balso forest" Then
  Call DeepBalsoForest
   Exit Sub
  End If
'-------------------
  
'--Slohovan Shore
 If FLocation = "slohovan shore" Then
  Call SlohovanShore
   Exit Sub
  End If
'----------------

'--Slohovan Cape
 If FLocation = "slohovan cape" Then
  If PeasantJacobQuestInFight = False Then
  Call SlohovanCape
   Exit Sub
  End If
 End If
'---------------

'--Port Alura
 If FLocation = "port alura" Then
  Call PortAlura
   Exit Sub
  End If
'------------

'--Port Glohova
 If FLocation = "port glohova" Then
  Call PortGlohova
   Exit Sub
 End If
 
 '--Vile Swamp
  If FLocation = "vile swamp" Then
   Call VileSwamp
    Exit Sub
  End If
 
 '--Village Of Glohova
  If FLocation = "village of glohova" Then
   Call VillageOfGlohova
    Exit Sub
  End If
  
 '--Glohova Cemetary
  If FLocation = "glohova cemetary" Then
   Call GlohovaCemetary
    Exit Sub
  End If
  
 '--Vile Lake
  If FLocation = "vile lake" Then
   Call VileLake
    Exit Sub
  End If

End Sub


Public Sub BalsoKingTalk()

If BalsoStartRoom <> 0 Then

 Call SetCompass("", "Village Of Balso", "", "")

  
  Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
   Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
    Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
     Call AddText("", BRIGHTCYAN, 8, False, False, False)
      Call AddText("Talk To King Balso", CYAN, 8, False, False, False)
       Call AddText("", BRIGHTCYAN, 8, False, False, False)
        Call AddText("==========", BRIGHTCYAN, 8, False, True, False)

'Wait For Input During Refugee

 Do While BalsoStartRoom = 1
  'If they leave, stop the do loop
    If FLocation <> "balso throne room" Then
      Exit Sub
       End If

  'If they type talk to king balso then talk back
    If PText = "talk to king balso" Then
    If BalsoRefugeeFight <> 1 Then
     Call AddText("King Balso: Why are you still here? Go kill those bandits terrorizing the camp!", BRIGHTBLUE, 9, False, False, False)
      Call NullPText
       Exit Sub
        End If
        End If
        
    If PText = "talk to king balso" Then
    If BalsoRefugeeFight = 1 Then
    If BalsoPortAlura = 0 Then
     Call AddText("King Balso: Great Job " & CName & "! My soldiers told me all about your victory at the refugee camp!", BRIGHTBLUE, 9, False, False, False)
      Call AddText("The Kings Advisor Rushes In And Begins Talking To The King...", BLUE, 9, True, False, False)
       Call AddText("Minutes Later Their Conversation Ends...", BLUE, 9, True, False, False)
        Call AddText("King Balso: I've been informed that there has been a string of strange unnatural events happening on the mainland...", BRIGHTBLUE, 9, False, False, False)
         Call AddText("King Balso: Since most of my soldiers are injured from the raid on the camp, I would like you to go investigate what is causing these event..", BRIGHTBLUE, 9, False, False, False)
          Call AddText("King Balso: I will have one of my personal ships and crew take you to the mainland from Port Alura", BRIGHTBLUE, 9, False, True, False)
           Call AddText("King Balso: Now off to the port with you!", BRIGHTBLUE, 9, False, False, False)
            Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "BalsoPortAlura", "1")
             BalsoPortAlura = 1
              Call NullPText
               Exit Sub
           End If
            End If
             End If
    
    If PText = "talk to king balso" Then
    If BalsoRefugeeFight = 1 Then
    If BalsoPortAlura = 1 Then
        Call AddText("King Balso: Shouldn't you be heading to the port?", BRIGHTBLUE, 9, False, False, False)
         Call NullPText
          Exit Sub
         End If
          End If
           End If
        
    DoEvents
   Loop
   


 
End If

End Sub

Public Sub BalsoRefugeeCamp()

Dim a As Boolean
Dim b As Boolean
Dim c As Boolean


a = False
b = True
c = False

Call SetCompass("", "Balso Forest", "", "")


If BalsoRefugeeFight = 0 Then
 BalsoRefugeeInFight = True
  Compass = False
   
 Call AddText("Injured Refugee: Ahh!!!!! Someone help us!!!!!", BRIGHTBLUE, 9, False, False, False)
  Call AddText("A Bandit attacks the injured refugee", BLUE, 9, True, False, False)
   Call AddText("Bandit Chief: Take the children!! And the gold!! Kill anyone who tries to stop you!!", BRIGHTBLUE, 9, False, False, False)
    Call AddText("A Bandit Rushes Towards You And Attacks!", BLUE, 9, True, False, False)
     
      Call SingleCreature("bandit", False)
    
'Loop until killed creature
Do While a = False
  
'initializes when bandit dies
 If Compass = True Then
  Compass = False
   Call AddText("Another Bandit Rushes Towards You And Attacks!", BLUE, 9, True, False, True)
    a = True
     b = False
      Call SingleCreature("bandit", False)
       End If

  If Health <= 0 Then
  Call ManualDeathCheck
   a = False
    b = True
     c = False
  End If
  
  DoEvents
 Loop
 
 
 
 
 'Loop until bandit 2 dies
Do While b = False
   
  'Initialize when bandit dies
  If b = False Then
   If Compass = True Then
    Compass = False
     Call AddText("Bandit Chief: Who are you? You dare kill my minions?", BRIGHTBLUE, 9, False, False, False)
      Call AddText("Bandit Chief: Hmph, a little pest like you deserves a proper beating!", BRIGHTBLUE, 9, False, False, False)
       Call AddText("The Bandit Chief Rushes Towards You", BLUE, 9, True, False, True)
        Call SingleCreature("Bandit Chief", False)
         b = True
            End If
             End If
  
  If Health <= 0 Then
   Call ManualDeathCheck
   a = False
    b = True
     c = False
      End If
  
  DoEvents
Loop

'loop until bandit chief dies
Do While b = True
 
 If Compass = True Then
 b = False
  Call AddText("Injured Soldier: The enemy is running!! Whoever you are, you've done a fine job!!", BRIGHTBLUE, 9, False, False, False)
   Call AddText("Injured Soldier: I'll be sure to let the king know about this!", BRIGHTBLUE, 9, False, False, False)
    Call AddText("Injured Soldier: I'm sure he will reward you!", BRIGHTBLUE, 9, False, False, False)
     Call AddText("Injured Soldier: You should go talk to the king as soon as possible!", BRIGHTBLUE, 9, False, False, False)
      Call AddText("The Injured Soldier Walks Back To Balso Castle", BLUE, 9, True, False, True)
       Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "BalsoRefugeeFight", "1")
        Call AddText("", BRIGHTBLUE, 9, False, False, False)
         BalsoRefugeeFight = 1
 End If

  If Health <= 0 Then
  Call ManualDeathCheck
   a = False
    b = True
     c = False
      End If

DoEvents
Loop

End If
'---------------------------------------------------
'------------End of Refugee Fight-------------------
'---------------------------------------------------
Dim randomnum As Integer


If BalsoRefugeeFight <> 0 Then

Call SetCompass("", "Balso Forest", "", "")
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To Refugees / Talk To Peasant Jacob", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)


AreaLoop:
'Talk To Refugees
Do While c = False
 
 If FLocation <> "balso refugee camp" Then
  Exit Sub
 End If
 
 
 If PText = "talk to refugees" Then
  Call NullPText
  
 randomnum = Rand(1, 4)

  If randomnum = 1 Then
   Call AddText("Female Refugee: I washed up on the shore a few days ago, I can't remember anything besides a giant sea creature attacking our ship..", BRIGHTBLUE, 9, False, False, False)
  End If

  If randomnum = 2 Then
   Call AddText("Child Refugee: Where's feefo?", BRIGHTBLUE, 9, False, False, False)
  End If
  
  If randomnum = 3 Then
   Call AddText("Random Refugee: I heard the bandit's around here took Jacob's younger brother in the raid..", BRIGHTBLUE, 9, False, False, False)
  End If
  
  If randomnum = 4 Then
   Call AddText("Male Refugee: ZzzzZZzZZZzzZzz...", BRIGHTBLUE, 9, False, False, False)
  End If
End If
  
 'talk to jacob for quest
  If PText = "talk to peasant jacob" Then
   
   
  If PeasantJacobQuest = 0 Then
   Call NullPText
   Call AddText("Peasant Jacob: Oh where has my little brother gone to...", BRIGHTBLUE, 9, False, False, False)
    Call AddText("Peasant Jacob Turns Towards You And Begins To Talk", BLUE, 9, True, False, False)
     Call AddText("Aren't you one of the kings warriors? I saw you fend off the bandits during the raid on our camp.", BRIGHTBLUE, 9, False, False, False)
      Call AddText("Do you think you could go find my little brother for me? (Y/N)", BRIGHTBLUE, 9, False, True, False)
        
        Do While c = False
         
         If FLocation <> "balso refugee camp" Then
          Exit Sub
         End If
        
    If PText = "y" Then
     Call NullPText
      Call AddText("Peasant Jacob: Oh thank you! Please find him as soon as possible!", BRIGHTBLUE, 9, False, False, False)
     Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "PeasantJacobQuest", "1")
    PeasantJacobQuest = 1
   GoTo AreaLoop
   End If
        
   If PText = "n" Then
    Call NullPText
     Call AddText("Peasant Jacob: Oh... Alright... Well if you change your mind come back to me any time", BRIGHTBLUE, 9, False, False, False)
      GoTo AreaLoop
   End If
  
    DoEvents
  Loop
  
  End If
 End If
  
  'talk to jacob while in quest
  
  If PText = "talk to peasant jacob" Then
   If PeasantJacobQuest = 1 Then
    Call NullPText
     Call AddText("Peasant Jacob: Still looking for my younger brother huh?", BRIGHTBLUE, 9, False, False, False)
    End If
   End If
  
  If PText = "talk to peasant jacob" Then
   If PeasantJacobQuest = 2 Then
    Call NullPText
     Call AddText("Peasant Jacob: Thank you so much for helping my younger brother!", BRIGHTBLUE, 9, False, False, False)
      Call AddText("Peasant Jacob: He told me all about your heroic efforts, the bandit that took him, everything!", BRIGHTBLUE, 9, False, False, False)
       Call AddText("Peasant Jacob: This is all I have to reward you with, so here, please take this for your kind deed!", BRIGHTBLUE, 9, False, False, False)
        Call CompletedQuest(50, 50)
         Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "PeasantJacobQuest", "3")
        End If
       End If
   
  If PText = "talk to peasant jacob" Then
   If PeasantJacobQuest = 3 Then
    Call NullPText
     Call AddText("Peasant Jacob: I still owe you! Thank you so much for helping my brother.", BRIGHTBLUE, 9, False, False, False)
      End If
       End If
       
 DoEvents
Loop





End If

End Sub

Public Sub DeepBalsoForest()

Dim a As Boolean

a = False

Call SetCompass("Balso Forest", "Slohovan Shore", "", "")
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)


'hunt command
Do While a = False

If FLocation <> "deep balso forest" Then
 Exit Sub
End If

If PText = "hunt" Then
 Call Hunt("bear cub", "tree dwelling goblin", "bandit", "wolf", "bear cub", "tree dwelling goblin", "bandit", "wolf", "tree dwelling goblin")
 Exit Sub
End If

DoEvents
Loop
'--------

End Sub

Public Sub SlohovanShore()
Dim a As Boolean
a = False

Call SetCompass("Deep Balso Forest", "Port Alura", "", "Slohovan Cape")
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)


'hunt command
Do While a = False

If FLocation <> "slohovan shore" Then
 Exit Sub
End If

If PText = "hunt" Then
 Call Hunt("red tongue salamander", "giant toad", "snapping turtle", "bandit", "red tongue salamander", "giant toad", "snapping turtle", "bandit", "giant snapping turtle")
 Exit Sub
End If

DoEvents
Loop


End Sub

Public Sub SlohovanCape()

Dim a As Boolean
a = False


Call SetCompass("", "", "Slohovan Shore", "")

'---------------JACOB QUEST---------------

If PeasantJacobQuest = 1 Then
 PeasantJacobQuestInFight = True
  Compass = False
   
 Call AddText("Young Boy: Help me!! Someone!! Anyone!!", BRIGHTBLUE, 9, False, False, False)
  Call AddText("Bandit Captain: Shut up!! No one is going to help you!!", BRIGHTBLUE, 9, False, False, False)
   Call AddText("You Walk Closer Towards The Bandit And Boy", BLUE, 9, True, False, False)
    Call AddText("Bandit Captain: Huh?? Who are you?? Oh you want the boy huh? Well, you'll have to take him!", BRIGHTBLUE, 9, False, False, False)
     Call AddText("The Bandit Captain Charges Towards You", BLUE, 9, True, False, False)
      Call SingleCreature("bandit captain", False)
    
'Loop until killed creature
Do While a = False
  
'initializes when bandit dies
 If Compass = True Then
  If FLocation = "slohovan cape" Then
  PeasantJacobQuestInFight = False
   Call AddText("Bandit Captain: Ugh... I know when i'm beaten, take the boy..", BRIGHTBLUE, 9, False, False, True)
    Call AddText("The Bandit Captain Falls To The Ground", BLUE, 9, True, False, False)
     Call AddText("Young Boy: Oh thank you!! I owe you my life!! Well, I better return to the camp now..", BRIGHTBLUE, 9, False, False, False)
      Call AddText("The Boy Walk Off Towards Balso Forest", BLUE, 9, True, False, False)
       FLocation = "slohovan shore"
        Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "PeasantJacobQuest", "2")
         PeasantJacobQuest = 2
    a = True
     Exit Sub
       End If
        End If
  
   If Health <= 0 Then
    Call ManualDeathCheck
     Exit Sub
      End If
     
  DoEvents
 Loop



End If



'================================
If PeasantJacobQuest <> 1 Then
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)

'hunt command
Do While a = False

If FLocation <> "slohovan cape" Then
 Exit Sub
End If

If PText = "hunt" Then
 If PeasantJacobQuest <> 1 Then
 Call Hunt("red tongue salamander", "giant toad", "snapping turtle", "bandit", "red tongue salamander", "giant toad", "snapping turtle", "bandit", "giant snapping turtle")
  Exit Sub
 End If
End If

DoEvents
Loop

End If
'==




End Sub

Public Sub PortAlura()

Dim SailorRandomTalk As Integer

Dim a As Boolean
a = False


Call SetCompass("Slohovan Shore", "", "", "")
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To Sailors / Talk To Kings Sailors", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)



TalkLoop:

If FLocation <> "port alura" Then
 Exit Sub
End If



Do While a = False

'--talk to sailors
 If PText = "talk to sailors" Then
  
  SailorRandomTalk = Rand(1, 3)
   
   If SailorRandomTalk = 1 Then
    Call AddText("1st Mate: The sea is a strange place, it seems endless...", BRIGHTBLUE, 9, False, False, False)
    Call NullPText
     End If
   
   If SailorRandomTalk = 2 Then
    Call AddText("Pirate: We pirates and sailor mostly get along.. mostly.. heh..", BRIGHTBLUE, 9, False, False, False)
    Call NullPText
     End If
     
   If SailorRandomTalk = 3 Then
    Call AddText("Sailor: I wonder how many ports their are in the world? hmm.. It makes me wonder..", BRIGHTBLUE, 9, False, False, False)
    Call NullPText
   End If

 End If
'---------------



'--talk to kings sailors


If PText = "talk to kings sailors" Then

If BalsoPortAlura = 0 Then
 Call AddText("King's Sailor: What do you want? It would be best for you to go on your way..", BRIGHTBLUE, 9, False, False, False)
  Call NullPText
   GoTo TalkLoop
End If

If BalsoPortAlura = 1 Then
 Call AddText("King's Sailor: The king told me that you need to get to the mainland?", BRIGHTBLUE, 9, False, False, False)
  Call AddText("King's Sailor: Well, we are the king's personal crew after all, dont worry we'll get you there.", BRIGHTBLUE, 9, False, False, False)
   Call AddText("King's Sailor: Just remember, you can return here anytime from Port Glohova.", BRIGHTBLUE, 9, False, False, False)
    Call AddText("King's Sailor: So, would you like to sail to Port Glohova? (Y/N)", BRIGHTBLUE, 9, False, True, False)
     End If
    
Do While a = False
 
 'if yes---
 If PText = "y" Then
    Call NullPText
     Call AddText("King's Sailor: Alright, let's set sail for Port Glohova!", BRIGHTBLUE, 9, False, False, False)
      Call SetLocation("Port Glohova")
       Location = "Port Glohova"
        FLocation = "port glohova"
         Call AddText("You Are Now In Port Glohova", BROWN, 9, False, True, True)
          Call CheckEvents
         End If
          
  '---------
  
  
  'if no---
 If PText = "n" Then
  Call NullPText
   Call AddText("King's Sailor: Alright then, just remember, anytime you want to go to the mainland, come to me!", BRIGHTBLUE, 9, False, False, False)
    GoTo TalkLoop
    End If
   '-------

 DoEvents
Loop

End If

 DoEvents
Loop
 
 
End Sub

Public Sub ManualDeathCheck()

'Check for player death
 If Health <= 0 Then
 Call DoneBattle
  Call AddText("=====", RED, 9, False, False, False)
   Call AddText("You Have Died!!", RED, 9, False, False, True)
    Call AddText("=====", RED, 9, False, False, False)
     Health = MaxHealth
      
      'Subtract The Experience Penalty + Display It
       Call DeathExperiencePenalty
     
      Call SetLocation(ReviveLocation)
       Call UpdateCompass
        Location = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "Location")
         FLocation = LCase(Location)
          Call CheckEvents
           Call SaveGame
           End If

End Sub

Public Sub PortGlohova()

Dim a As Boolean
a = False


Call SetCompass("Village Of Glohova", "Vile Swamp", "", "")

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To Kings Sailors", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)

If FLocation <> "port glohova" Then
 Exit Sub
End If

Do While a = False

TalkLoop:

If PText = "talk to kings sailors" Then
 Call NullPText
  Call AddText("King's Sailor: Ready to head back to Port Alura? (Y/N)", BRIGHTBLUE, 9, False, False, False)
  
  Do While a = False
  
  
 If PText = "y" Then
  Call NullPText
   Call AddText("King's Sailor: Let's set sail to Port Alura!", BRIGHTBLUE, 9, False, False, False)
    Call SetLocation("Port Alura")
     Location = "Port Alura"
      FLocation = "port alura"
       Call AddText("You Are Now In Port Alura", BROWN, 9, False, True, True)
        Call CheckEvents
 
 End If
 
 
 
 If PText = "n" Then
  Call NullPText
   Call AddText("King's Sailor: Alright, If you change your mind we'll be here!", BRIGHTBLUE, 9, False, False, False)
    GoTo TalkLoop
 End If
  
  
 DoEvents
Loop


End If


 DoEvents
Loop


End Sub

Public Sub VileSwamp()

Dim a As Boolean
a = False

Call SetCompass("Port Glohova", "Vile Lake", "", "")

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)

'hunt command
Do While a = False

If FLocation <> "vile swamp" Then
 Exit Sub
End If

If PText = "hunt" Then
 Call Hunt("Diseased Toad", "Viper", "Giant Spider", "Diseased Salamander", "Diseased Toad", "Viper", "Giant Spider", "Diseased Salamander", "Mother Spider")
  Exit Sub
End If

DoEvents
Loop
'--------




End Sub

Public Sub VillageOfGlohova()

Dim VillageRandTalk As Integer
Dim a As Boolean


a = False

Call SetCompass("", "Port Glohova", "", "Glohova Cemetary")

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To Villagers / Shop / Use Ressurection Shrine", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To King Balsos Assistant / Talk To Witchdoctor", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)



AreaLoop:

Do While a = False

If FLocation <> "village of glohova" Then
 Exit Sub
End If


'Use Ressurection Shrine--
 If PText = "use ressurection shrine" Then
  Call SetReviveLocation("Village Of Glohova", 200)
   End If
 '------------------------

 
 'Talk to villagers-------
  If PText = "talk to villagers" Then
   Call NullPText
   
    VillageRandTalk = Rand(1, 4)
   
   
   If VillageRandTalk = 1 Then
    Call AddText("Villager: I wouldn't go south of Port Glohova without posion resistant armor..", BRIGHTBLUE, 9, False, False, False)
   End If
   
   If VillageRandTalk = 2 Then
    Call AddText("Young Boy: A while ago I got lost in the wilderness, I didn't know where I was, After a while I ran into a giant frog, he was incredibly large! I had never seen anything like it before...", BRIGHTBLUE, 9, False, False, False)
   End If
   
   If VillageRandTalk = 3 Then
    Call AddText("Old Man: Back in my day, I was a fine warrior, One thing I learned was that some armors can repel certain enemy effects! Like posion!", BRIGHTBLUE, 9, False, False, False)
   End If
   
   If VillageRandTalk = 4 Then
    Call AddText("Traveller: Hello there, I'm visiting from the Kingdom Of Zen, also known as the capital!", BRIGHTBLUE, 9, False, False, False)
   End If
  
 End If
  '-----------------------
   
  'Talk To King Balsos Assistant
  
   If PText = "talk to king balsos assistant" Then
    Call NullPText
     
    'If the variable isnt in .eoc add it---
     If EventExist("VillageOfGlohovaKingAssistant") = False Then
      Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "VillageOfGlohovaKingAssistant", "1")
       VillageOfGlohovaKingAssistant = 1
        End If
        '---------------------------
        
        'while in quest
         If VillageOfGlohovaKingAssistant = 2 Then
          Call AddText("King Balso's Assistant: Shouldn't you be investigating at the Vile Lake?..", BRIGHTBLUE, 9, False, False, False)
           End If
           '---------
        
       'initial quest
        If VillageOfGlohovaKingAssistant = 1 Then
         Call AddText("King Balso's Assistant: Oh good,  you made it safely across the sea.", BRIGHTBLUE, 9, False, False, False)
          Call AddText("King Balso's Assistant: Well we should get down to buisiness then shouldn't we..", BRIGHTBLUE, 9, False, False, False)
           Call AddText("King Balso's Assistant: The king wants you to find out whats causing the lands near here to wither.", BRIGHTBLUE, 9, False, False, False)
            Call AddText("King Balso's Assistant: I heard that there's a cult that does hellbent rituals at the Vile Lake.", BRIGHTBLUE, 9, False, False, False)
             Call AddText("King Balso's Assistant: Maybe those rituals are linked to the dieing land?", BRIGHTBLUE, 9, False, False, False)
              Call AddText("King Balso's Assistant: Will you go check out the Vile Lake? (Y/N)", BRIGHTBLUE, 9, False, True, False)
                    
             Do While a = False
              
              If PText = "y" Then
               Call NullPText
                Call AddText("King Balso's Assistant: Excellent! Come back to me when you find out whats going on!", BRIGHTBLUE, 9, False, False, False)
                 Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "VillageOfGlohovaKingAssistant", "2")
                  VillageOfGlohovaKingAssistant = 2
                   GoTo AreaLoop
                   End If
                   
              If PText = "n" Then
               Call NullPText
                Call AddText("King Balso's Assistant: Well, come back anytime you'd like to investigate..", BRIGHTBLUE, 9, False, False, False)
                 GoTo AreaLoop
                  End If
                 
               DoEvents
                Loop
                
        End If
      
        '-------------

      'After Quest
       If VillageOfGlohovaKingAssistant = 4 Then
        Call AddText("King Balso's Assistant: May your adventures bring new knowledge and self discovery!", BRIGHTBLUE, 9, False, False, False)
         End If

      'End of quest
        If VillageOfGlohovaKingAssistant = 3 Then
         Call AddText("King Balso's Assistant: Did you find out if the cult is linked to what's happening?", BRIGHTBLUE, 9, False, False, False)
          Call AddText("King Balso's Assistant: No? hmmm...", BRIGHTBLUE, 9, False, False, False)
           Call AddText("The King's Assistant Begins To Think", BLUE, 9, True, False, False)
            Call AddText("King Balso's Assistant: Well, we have no leads on who's causing all of this..", BRIGHTBLUE, 9, False, False, False)
             Call AddText("King Balso's Assistant: So, I believe it would be best if you ventured out on your own for now.", BRIGHTBLUE, 9, False, True, False)
              Call AddText("King Balso's Assistant: While you travel you may find out more about what's going on.", BRIGHTBLUE, 9, False, False, False)
               Call AddText("King Balso's Assistant: You may even find out more about yourself! Who knows.", BRIGHTBLUE, 9, False, False, False)
                Call AddText("King Balso's Assistant: But for now, something's telling me you should investigate the desert while venturing..", BRIGHTBLUE, 9, False, True, False)
                 Call AddText("King Balso's Assistant: Oh! One last thing, heres your reward for investigating the cult!", BRIGHTBLUE, 9, False, False, False)
                  Call CompletedQuest(100, 150)
                   Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "VillageOfGlohovaKingAssistant", "4")
                    VillageOfGlohovaKingAssistant = 4
                   End If
                   
 
         
End If

'------------------------------
'------------------------------



'shop
 If PText = "shop" Then
  Call NullPText
   Call Shop("Potion", "Soul Potion", "Large Potion", "Large Soul Potion", "", "", "", "", "", "")
    End If
'------









 DoEvents
Loop



End Sub

Public Sub GlohovaCemetary()

Call SetCompass("", "", "Village Of Glohova", "")

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Talk To Crypt Keeper", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)


End Sub

Public Sub VileLake()
Dim a As Boolean
a = False


Call SetCompass("Vile Swamp", "", "Deep Vile Lake", "Vile Stream")

'--If Not In Quest--
If VillageOfGlohovaKingAssistant <> 2 Then

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)

'hunt command
Do While a = False

If FLocation <> "vile lake" Then
 Exit Sub
End If

If PText = "hunt" Then
 Call Hunt("Rotting Ent", "Viper", "Giant Spider", "Diseased Salamander", "Diseased Toad", "Viper", "Rotting Ent", "Diseased Salamander", "Mother Spider")
  Exit Sub
End If

DoEvents
Loop

End If

'--------------------



'--If IN Quest
If VillageOfGlohovaKingAssistant = 2 Then

Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("==Area Commands==", BRIGHTCYAN, 8, False, True, False)
Call AddText("==========", BRIGHTCYAN, 8, False, True, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("Hunt / Talk To Cult", CYAN, 8, False, False, False)
Call AddText("", BRIGHTCYAN, 8, False, False, False)
Call AddText("===========", BRIGHTCYAN, 8, False, True, False)

'hunt command
Do While a = False

If FLocation <> "vile lake" Then
 Exit Sub
End If


If PText = "hunt" Then
 Call Hunt("Rotting Ent", "Viper", "Giant Spider", "Diseased Salamander", "Diseased Toad", "Viper", "Rotting Ent", "Diseased Salamander", "Mother Spider")
  Exit Sub
End If


'talk to cult
If PText = "talk to cult" Then
Compass = False
 Call AddText("Cult Members: Huuuummmmmmmm.... Hummmmmmmmm........", BRIGHTBLUE, 9, False, False, False)
  Call AddText("The Cult Continues Chanting In A Circle", BLUE, 9, True, False, False)
   Call AddText("You Walk Closer...", BLUE, 9, True, False, False)
    Call AddText("Cult Leader: Ahhh, " & CName & "....", BRIGHTBLUE, 9, False, False, False)
     Call AddText("Cult Leader: The massssssster is not happy with you..", BRIGHTBLUE, 9, False, False, False)
      Call AddText("Cult Leader: You don't remember what hasssss happened do you?.. (Y/N)", BRIGHTBLUE, 9, False, True, False)
       
       
       Do While a = False
       
       
       
       
       
       If PText = "y" Then
        Call NullPText
       Call AddText("Cult Leader: I doubt you remember... But for now...", BRIGHTBLUE, 9, False, False, False)
        Call AddText("Cult Leader: I'm sssssssure the masssster will keep a closssse eye on you...", BRIGHTBLUE, 9, False, False, False)
         Call AddText("Cult Leader: Eventsssss of epic proportion will be taking place sssssooon...", BRIGHTBLUE, 9, False, True, False)
          Call AddText("Cult Leader: Now... We musssst take our leave..", BRIGHTBLUE, 9, False, False, False)
           Call AddText("Cult Leader: There are sssssome important matters that need attending...", BRIGHTBLUE, 9, False, False, False)
            Call AddText("The Cult Members Vanish Into A Flash Of Darkness", BLUE, 9, True, False, False)
             Compass = True
      Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "VillageOfGlohovaKingAssistant", "3")
       VillageOfGlohovaKingAssistant = 3
        Call CheckEvents
                 
       End If
       
       
       
       
       
       
       If PText = "n" Then
         Call NullPText
        Call AddText("Cult Leader: I sssssee, All shall be revealed in due time..", BRIGHTBLUE, 9, False, False, False)
         Call AddText("Cult Leader: I'm sssssssure the masssster will keep a closssse eye on you from now on...", BRIGHTBLUE, 9, False, False, False)
          Call AddText("Cult Leader: Eventsssss of epic proportion will be taking place sssssooon...", BRIGHTBLUE, 9, False, True, False)
           Call AddText("Cult Leader: But for now.. We musssst take our leave..", BRIGHTBLUE, 9, False, False, False)
            Call AddText("Cult Leader: There are sssssome important matters that need attending...", BRIGHTBLUE, 9, False, False, False)
             Call AddText("The Cult Members Vanish Into A Flash Of Darkness", BLUE, 9, True, False, False)
              Compass = True
      Call PutVar((App.Path & "\Characters\" & CName & "Events" & ".EoC"), "Events", "VillageOfGlohovaKingAssistant", "3")
       VillageOfGlohovaKingAssistant = 3
        Call CheckEvents
       
       End If
       
       
       
    DoEvents
     Loop

End If

DoEvents
Loop


End If


End Sub
