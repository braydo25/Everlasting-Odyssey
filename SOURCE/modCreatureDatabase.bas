Attribute VB_Name = "modCreatureDatabase"
Option Explicit

Public Sub SingleCreature(Creature As String, CanRun As Boolean)

Creature = LCase(Creature)



'--- Bandit ---
If Creature = "bandit" Then
  
  CreatureName = "Bandit"
   CreatureHealth = 40
    CreatureStrength = 12
     CreatureDefence = 1
      CreatureExperience = 5
       CreatureGold = 5
        CreaturePosion = False
    Call FullBattle(CanRun)
    
End If

'--- Wolf ---
If Creature = "wolf" Then
 
 CreatureName = "Wolf"
  CreatureHealth = 25
   CreatureStrength = 10
    CreatureDefence = 2
     CreatureExperience = 3
      CreatureGold = 2
       CreaturePosion = False
    Call FullBattle(CanRun)
End If

'--- Ferocious Rat ---
If Creature = "ferocious rat" Then
  
 CreatureName = "Ferocious Rat"
  CreatureHealth = 15
   CreatureStrength = 9
    CreatureDefence = 5
     CreatureExperience = 4
      CreatureGold = 3
       CreaturePosion = False
    Call FullBattle(CanRun)
End If
    
 '--- Treetop Spitter ---
 If Creature = "treetop spitter" Then
  
  CreatureName = "Treetop Spitter"
   CreatureHealth = 55
    CreatureStrength = 12
     CreatureDefence = 10
      CreatureExperience = 9
       CreatureGold = 4
        CreaturePosion = False
    Call FullBattle(CanRun)
End If

 '--- Bandit Chief ---
 If Creature = "bandit chief" Then
 
  CreatureName = "Bandit Chief"
   CreatureHealth = 105
    CreatureStrength = 13
     CreatureDefence = 3
      CreatureExperience = 20
       CreatureGold = 30
        CreaturePosion = False
    Call FullBattle(CanRun)
End If

 '--- Bear Cub ---
 If Creature = "bear cub" Then
 
 CreatureName = "Bear Cub"
  CreatureHealth = 35
   CreatureStrength = 11
    CreatureDefence = 12
     CreatureExperience = 10
      CreatureGold = 2
       CreaturePosion = False
     Call FullBattle(CanRun)
End If

 '--- Tree Dwelling Goblin ---
 If Creature = "tree dwelling goblin" Then
  
  CreatureName = "Tree Dwelling Globlin"
   CreatureHealth = 20
    CreatureStrength = 15
     CreatureDefence = 2
      CreatureExperience = 6
       CreatureGold = 10
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
  
 '---Red Tongue Salamander---
 If Creature = "red tongue salamander" Then

  CreatureName = "Red Tongue Salamander"
   CreatureHealth = 60
    CreatureStrength = 14
     CreatureDefence = 6
      CreatureExperience = 10
       CreatureGold = 5
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
    
 '--- Giant Toad ---
 If Creature = "giant toad" Then
 
  CreatureName = "Giant Toad"
   CreatureHealth = 110
    CreatureStrength = 8
     CreatureDefence = 15
      CreatureExperience = 15
       CreatureGold = 7
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
     
 '--- Snapping Turtle ---
 If Creature = "snapping turtle" Then

  CreatureName = "Snapping Turtle"
   CreatureHealth = 70
    CreatureStrength = 9
     CreatureDefence = 16
      CreatureExperience = 18
       CreatureGold = 9
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
      
 '--- Bandit Captain ---
 If Creature = "bandit captain" Then
 
  CreatureName = "Bandit Captain"
   CreatureHealth = 180
    CreatureStrength = 20
     CreatureDefence = 16
      CreatureExperience = 40
       CreatureGold = 40
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
    
 '--- Diseased Toad ---
 If Creature = "diseased toad" Then
 
  CreatureName = "Diseased Toad"
   CreatureHealth = 160
    CreatureStrength = 18
     CreatureDefence = 10
      CreatureExperience = 35
       CreatureGold = 30
        CreaturePosion = True
      Call FullBattle(CanRun)
End If

 '--- Viper ---
 If Creature = "viper" Then
 
  CreatureName = "Viper"
   CreatureHealth = 120
    CreatureStrength = 17
     CreatureDefence = 10
      CreatureExperience = 20
       CreatureGold = 15
        CreaturePosion = True
      Call FullBattle(CanRun)
End If

 '--- Giant Spider ---
 If Creature = "giant spider" Then
  
  CreatureName = "Giant Spider"
   CreatureHealth = 195
    CreatureStrength = 22
     CreatureDefence = 14
      CreatureExperience = 45
       CreatureGold = 30
        CreaturePosion = True
      Call FullBattle(CanRun)
End If

 '--- Diseased Salamander ---
 If Creature = "diseased salamander" Then
 
  CreatureName = "Diseased Salamander"
   CreatureHealth = 160
    CreatureStrength = 19
     CreatureDefence = 21
      CreatureExperience = 37
       CreatureGold = 14
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
    
 '--- Rotting Ent ---
 If Creature = "rotting ent" Then
  
  CreatureName = "Rotting Ent"
   CreatureHealth = 185
    CreatureStrength = 24
     CreatureDefence = 19
      CreatureExperience = 50
       CreatureGold = 15
        CreaturePosion = False
      Call FullBattle(CanRun)
End If
    
    
End Sub

Public Sub MultiCreature1Database(MultiCreat1Name As String)


'--- Bandit ---
If MultiCreat1Name = "bandit" Then
  
  MultiCreature1Name = "Bandit"
   MultiCreature1Health = 40
    MultiCreature1Strength = 12
     MultiCreature1Defence = 1
      MultiCreature1Experience = 5
       MultiCreature1Gold = 5
        MultiCreature1Posion = False

End If

'--- Wolf ---
If MultiCreat1Name = "wolf" Then
 
 MultiCreature1Name = "Wolf"
  MultiCreature1Health = 25
   MultiCreature1Strength = 10
    MultiCreature1Defence = 2
     MultiCreature1Experience = 3
      MultiCreature1Gold = 2
       MultiCreature1Posion = False

End If

'--- Ferocious Rat ---
If MultiCreat1Name = "ferocious rat" Then
  
 MultiCreature1Name = "Ferocious Rat"
  MultiCreature1Health = 15
   MultiCreature1Strength = 9
    MultiCreature1Defence = 5
     MultiCreature1Experience = 4
      MultiCreature1Gold = 3
       MultiCreature1Posion = False

End If
    
 '--- Treetop Spitter ---
 If MultiCreat1Name = "treetop spitter" Then
  
  MultiCreature1Name = "Treetop Spitter"
   MultiCreature1Health = 55
    MultiCreature1Strength = 12
     MultiCreature1Defence = 4
      MultiCreature1Experience = 9
       MultiCreature1Gold = 10
        MultiCreature1Posion = False

End If

 '--- Bandit Chief ---
 If MultiCreat1Name = "bandit chief" Then
 
  MultiCreature1Name = "Bandit Chief"
   MultiCreature1Health = 105
    MultiCreature1Strength = 13
     MultiCreature1Defence = 3
      MultiCreature1Experience = 20
       MultiCreature1Gold = 30
        MultiCreature1Posion = False
 
End If

 '--- Bear Cub ---
 If MultiCreat1Name = "bear cub" Then
 
 MultiCreature1Name = "Bear Cub"
  MultiCreature1Health = 35
   MultiCreature1Strength = 11
    MultiCreature1Defence = 12
     MultiCreature1Experience = 10
      MultiCreature1Gold = 7
       MultiCreature1Posion = False
      
End If

 '--- Tree Dwelling Goblin ---
 If MultiCreat1Name = "tree dwelling goblin" Then
  
  MultiCreature1Name = "Tree Dwelling Globlin"
   MultiCreature1Health = 20
    MultiCreature1Strength = 15
     MultiCreature1Defence = 2
      MultiCreature1Experience = 6
       MultiCreature1Gold = 10
        MultiCreature1Posion = False
       
End If

 '---Red Tongue Salamander---
 If MultiCreat1Name = "red tongue salamander" Then

  MultiCreature1Name = "Red Tongue Salamander"
   MultiCreature1Health = 60
    MultiCreature1Strength = 14
     MultiCreature1Defence = 6
      MultiCreature1Experience = 10
       MultiCreature1Gold = 5
        MultiCreature1Posion = False

End If

 '---Giant Toad---
 If MultiCreat1Name = "giant toad" Then
 
  MultiCreature1Name = "Giant Toad"
   MultiCreature1Health = 110
    MultiCreature1Strength = 8
     MultiCreature1Defence = 15
      MultiCreature1Experience = 15
       MultiCreature1Gold = 7
        MultiCreature1Posion = False

End If
 
 '---Snapping Turtle---
 If MultiCreat1Name = "snapping turtle" Then

  MultiCreature1Name = "Snapping Turtle"
   MultiCreature1Health = 70
    MultiCreature1Strength = 9
     MultiCreature1Defence = 16
      MultiCreature1Experience = 18
       MultiCreature1Gold = 9
        MultiCreature1Posion = False

End If

'---Bandit Captain---
 If MultiCreat1Name = "bandit captain" Then
 
  MultiCreature1Name = "Bandit Captain"
   MultiCreature1Health = 180
    MultiCreature1Strength = 20
     MultiCreature1Defence = 16
      MultiCreature1Experience = 40
       MultiCreature1Gold = 40
        MultiCreature1Posion = False
       
End If

 '---Diseased Toad---
 If MultiCreat1Name = "diseased toad" Then
 
  MultiCreature1Name = "Diseased Toad"
   MultiCreature1Health = 160
    MultiCreature1Strength = 15
     MultiCreature1Defence = 10
      MultiCreature1Experience = 35
       MultiCreature1Gold = 30
        MultiCreature1Posion = True
      
End If

 '---Viper---
 If MultiCreat1Name = "viper" Then
 
  MultiCreature1Name = "Viper"
   MultiCreature1Health = 120
    MultiCreature1Strength = 17
     MultiCreature1Defence = 10
      MultiCreature1Experience = 20
       MultiCreature1Gold = 15
        MultiCreature1Posion = True
    
End If

 '---Giant Spider---
 If MultiCreat1Name = "giant spider" Then
  
  MultiCreature1Name = "Giant Spider"
   MultiCreature1Health = 195
    MultiCreature1Strength = 22
     MultiCreature1Defence = 14
      MultiCreature1Experience = 45
       MultiCreature1Gold = 30
        MultiCreature1Posion = True

End If

 '--- Diseased Salamander ---
 If MultiCreat1Name = "diseased salamander" Then
 
  MultiCreature1Name = "Diseased Salamander"
   MultiCreature1Health = 160
    MultiCreature1Strength = 19
     MultiCreature1Defence = 21
      MultiCreature1Experience = 37
       MultiCreature1Gold = 14
        MultiCreature1Posion = False
    
End If

 '--- Rotting Ent ---
 If MultiCreat1Name = "rotting ent" Then
  
  MultiCreature1Name = "Rotting Ent"
   MultiCreature1Health = 185
    MultiCreature1Strength = 24
     MultiCreature1Defence = 19
      MultiCreature1Experience = 50
       MultiCreature1Gold = 15
        MultiCreature1Posion = False
      
End If

End Sub

Public Sub MultiCreature2Database(MultiCreat2Name As String)

'--- Bandit ---
If MultiCreat2Name = "bandit" Then
  
  MultiCreature2Name = "Bandit"
   MultiCreature2Health = 40
    MultiCreature2Strength = 12
     MultiCreature2Defence = 1
      MultiCreature2Experience = 5
       MultiCreature2Gold = 5
        MultiCreature2Posion = False

End If

'--- Wolf ---
If MultiCreat2Name = "wolf" Then
 
 MultiCreature2Name = "Wolf"
  MultiCreature2Health = 25
   MultiCreature2Strength = 10
    MultiCreature2Defence = 2
     MultiCreature2Experience = 3
      MultiCreature2Gold = 2
       MultiCreature2Posion = False

End If

'--- Ferocious Rat ---
If MultiCreat2Name = "ferocious rat" Then
  
 MultiCreature2Name = "Ferocious Rat"
  MultiCreature2Health = 15
   MultiCreature2Strength = 9
    MultiCreature2Defence = 5
     MultiCreature2Experience = 4
      MultiCreature2Gold = 3
       MultiCreature2Posion = False

End If
    
 '--- Treetop Spitter ---
 If MultiCreat2Name = "treetop spitter" Then
  
  MultiCreature2Name = "Treetop Spitter"
   MultiCreature2Health = 55
    MultiCreature2Strength = 12
     MultiCreature2Defence = 4
      MultiCreature2Experience = 9
       MultiCreature2Gold = 10
        MultiCreature2Posion = False

End If


 '--- Bandit Chief ---
 If MultiCreat2Name = "bandit chief" Then
 
  MultiCreature2Name = "Bandit Chief"
   MultiCreature2Health = 105
    MultiCreature2Strength = 13
     MultiCreature2Defence = 3
      MultiCreature2Experience = 20
       MultiCreature2Gold = 30
        MultiCreature2Posion = False
 
End If

 '--- Bear Cub ---
 If MultiCreat2Name = "bear cub" Then
 
 MultiCreature2Name = "Bear Cub"
  MultiCreature2Health = 35
   MultiCreature2Strength = 11
    MultiCreature2Defence = 12
     MultiCreature2Experience = 10
      MultiCreature2Gold = 7
       MultiCreature2Posion = False
      
End If
      
 '--- Tree Dwelling Goblin ---
 If MultiCreat2Name = "tree dwelling goblin" Then
  
  MultiCreature2Name = "Tree Dwelling Globlin"
   MultiCreature2Health = 20
    MultiCreature2Strength = 15
     MultiCreature2Defence = 2
      MultiCreature2Experience = 6
       MultiCreature2Gold = 10
        MultiCreature2Posion = False
       
End If

 '---Red Tongue Salamander---
 If MultiCreat2Name = "red tongue salamander" Then

  MultiCreature2Name = "Red Tongue Salamander"
   MultiCreature2Health = 60
    MultiCreature2Strength = 14
     MultiCreature2Defence = 6
      MultiCreature2Experience = 10
       MultiCreature2Gold = 5
        MultiCreature2Posion = False

End If

 '---Giant Toad---
 If MultiCreat2Name = "giant toad" Then
 
  MultiCreature2Name = "Giant Toad"
   MultiCreature2Health = 110
    MultiCreature2Strength = 8
     MultiCreature2Defence = 15
      MultiCreature2Experience = 15
       MultiCreature2Gold = 7
        MultiCreature2Posion = False

End If

'---Snapping Turtle---
 If MultiCreat2Name = "snapping turtle" Then

  MultiCreature2Name = "Snapping Turtle"
   MultiCreature2Health = 70
    MultiCreature2Strength = 9
     MultiCreature2Defence = 16
      MultiCreature2Experience = 18
       MultiCreature2Gold = 9
        MultiCreature2Posion = False
 
End If

'---Bandit Captain---
 If MultiCreat2Name = "bandit captain" Then
 
  MultiCreature2Name = "Bandit Captain"
   MultiCreature2Health = 180
    MultiCreature2Strength = 20
     MultiCreature2Defence = 16
      MultiCreature2Experience = 40
       MultiCreature2Gold = 40
        MultiCreature2Posion = False
       
End If

 '---Diseased Toad---
 If MultiCreat2Name = "diseased toad" Then
 
  MultiCreature2Name = "Diseased Toad"
   MultiCreature2Health = 160
    MultiCreature2Strength = 15
     MultiCreature2Defence = 10
      MultiCreature2Experience = 35
       MultiCreature2Gold = 30
        MultiCreature2Posion = True
      
End If

 '---Viper---
 If MultiCreat2Name = "viper" Then
 
  MultiCreature2Name = "Viper"
   MultiCreature2Health = 120
    MultiCreature2Strength = 17
     MultiCreature2Defence = 10
      MultiCreature2Experience = 20
       MultiCreature2Gold = 15
        MultiCreature2Posion = True
    
End If

 '---Giant Spider---
 If MultiCreat2Name = "giant spider" Then
  
  MultiCreature2Name = "Giant Spider"
   MultiCreature2Health = 195
    MultiCreature2Strength = 22
     MultiCreature2Defence = 14
      MultiCreature2Experience = 45
       MultiCreature2Gold = 30
        MultiCreature2Posion = True

End If

 '--- Diseased Salamander ---
 If MultiCreat2Name = "diseased salamander" Then
 
  MultiCreature2Name = "Diseased Salamander"
   MultiCreature2Health = 160
    MultiCreature2Strength = 19
     MultiCreature2Defence = 21
      MultiCreature2Experience = 37
       MultiCreature2Gold = 14
        MultiCreature2Posion = False
    
End If

 '--- Rotting Ent ---
 If MultiCreat2Name = "rotting ent" Then
  
  MultiCreature2Name = "Rotting Ent"
   MultiCreature2Health = 185
    MultiCreature2Strength = 24
     MultiCreature2Defence = 19
      MultiCreature2Experience = 50
       MultiCreature2Gold = 15
        MultiCreature2Posion = False
      
End If

End Sub


Public Sub BossCreatureDatabase(Creature As String, CanRun As Boolean)


Creature = LCase(Creature)



'--- Bandit ---
If Creature = "bandit" Then
  
  CreatureName = "Bandit"
   CreatureHealth = 40
    CreatureStrength = 12
     CreatureDefence = 1
      CreatureExperience = 5
       CreatureGold = 5
        CreaturePosion = False
    Call BossBattle(CanRun)
    
End If

'--- Wolf ---
If Creature = "wolf" Then
 
 CreatureName = "Wolf"
  CreatureHealth = 25
   CreatureStrength = 10
    CreatureDefence = 2
     CreatureExperience = 3
      CreatureGold = 2
       CreaturePosion = False
    Call BossBattle(CanRun)
End If

'--- Ferocious Rat ---
If Creature = "ferocious rat" Then
  
 CreatureName = "Ferocious Rat"
  CreatureHealth = 15
   CreatureStrength = 9
    CreatureDefence = 5
     CreatureExperience = 4
      CreatureGold = 3
       CreaturePosion = False
    Call BossBattle(CanRun)
End If
    
 '--- Treetop Spitter ---
 If Creature = "treetop spitter" Then
  
  CreatureName = "Treetop Spitter"
   CreatureHealth = 55
    CreatureStrength = 12
     CreatureDefence = 10
      CreatureExperience = 9
       CreatureGold = 4
        CreaturePosion = False
    Call BossBattle(CanRun)
End If

 '--- Bandit Chief ---
 If Creature = "bandit chief" Then
 
  CreatureName = "Bandit Chief"
   CreatureHealth = 105
    CreatureStrength = 13
     CreatureDefence = 3
      CreatureExperience = 20
       CreatureGold = 30
        CreaturePosion = False
    Call BossBattle(CanRun)
End If

 '--- Bear Cub ---
 If Creature = "bear cub" Then
 
 CreatureName = "Bear Cub"
  CreatureHealth = 35
   CreatureStrength = 11
    CreatureDefence = 12
     CreatureExperience = 10
      CreatureGold = 2
       CreaturePosion = False
     Call BossBattle(CanRun)
End If

 '--- Tree Dwelling Goblin ---
 If Creature = "tree dwelling goblin" Then
  
  CreatureName = "Tree Dwelling Globlin"
   CreatureHealth = 20
    CreatureStrength = 15
     CreatureDefence = 2
      CreatureExperience = 6
       CreatureGold = 10
        CreaturePosion = False
      Call BossBattle(CanRun)
End If
  
 '---Red Tongue Salamander---
 If Creature = "red tongue salamander" Then

  CreatureName = "Red Tongue Salamander"
   CreatureHealth = 60
    CreatureStrength = 14
     CreatureDefence = 6
      CreatureExperience = 10
       CreatureGold = 5
        CreaturePosion = False
      Call BossBattle(CanRun)
End If
    
 '---Giant Toad---
 If Creature = "giant toad" Then
 
  CreatureName = "Giant Toad"
   CreatureHealth = 110
    CreatureStrength = 8
     CreatureDefence = 15
      CreatureExperience = 15
       CreatureGold = 7
        CreaturePosion = False
      Call BossBattle(CanRun)
End If
     
 '---Snapping Turtle---
 If Creature = "snapping turtle" Then

  CreatureName = "Snapping Turtle"
   CreatureHealth = 70
    CreatureStrength = 9
     CreatureDefence = 16
      CreatureExperience = 18
       CreatureGold = 9
        CreaturePosion = False
      Call BossBattle(CanRun)
End If
      
 '---Giant Snapping Turtle (BOSS CREATURE)---
 If Creature = "giant snapping turtle" Then
   
  CreatureName = "Giant Snapping Turtle"
   CreatureHealth = 145
    CreatureStrength = 19
     CreatureDefence = 18
      CreatureExperience = 45
       CreatureGold = 20
        CreaturePosion = False
      Call BossBattle(CanRun)
End If

'---Bandit Captain---
 If Creature = "bandit captain" Then
 
  CreatureName = "Bandit Captain"
   CreatureHealth = 180
    CreatureStrength = 20
     CreatureDefence = 16
      CreatureExperience = 40
       CreatureGold = 40
        CreaturePosion = False
       Call BossBattle(CanRun)
End If

   
 '---Diseased Toad---
 If Creature = "diseased toad" Then
 
  CreatureName = "Diseased Toad"
   CreatureHealth = 160
    CreatureStrength = 15
     CreatureDefence = 10
      CreatureExperience = 35
       CreatureGold = 30
        CreaturePosion = True
      Call BossBattle(CanRun)
End If

'---Viper---
 If Creature = "viper" Then
 
  CreatureName = "Viper"
   CreatureHealth = 120
    CreatureStrength = 17
     CreatureDefence = 10
      CreatureExperience = 20
       CreatureGold = 15
        CreaturePosion = True
      Call BossBattle(CanRun)
End If

 '---Giant Spider---
 If Creature = "giant spider" Then
  
  CreatureName = "Giant Spider"
   CreatureHealth = 195
    CreatureStrength = 22
     CreatureDefence = 14
      CreatureExperience = 45
       CreatureGold = 30
        CreaturePosion = True
      Call BossBattle(CanRun)
End If

 '---Mother Spider(BOSS CREATURE)---
 If Creature = "mother spider" Then
  
  CreatureName = "Mother Spider"
   CreatureHealth = 235
    CreatureStrength = 29
     CreatureDefence = 20
      CreatureExperience = 80
       CreatureGold = 75
        CreaturePosion = True
      Call BossBattle(CanRun)
End If
  
 '--- Diseased Salamander ---
 If Creature = "diseased salamander" Then
 
  CreatureName = "Diseased Salamander"
   CreatureHealth = 160
    CreatureStrength = 19
     CreatureDefence = 21
      CreatureExperience = 37
       CreatureGold = 14
        CreaturePosion = False
      Call BossBattle(CanRun)
End If
  
 '--- Rotting Ent ---
 If Creature = "rotting ent" Then
  
  CreatureName = "Rotting Ent"
   CreatureHealth = 185
    CreatureStrength = 24
     CreatureDefence = 19
      CreatureExperience = 50
       CreatureGold = 15
        CreaturePosion = False
      Call BossBattle(CanRun)
End If
  
End Sub
