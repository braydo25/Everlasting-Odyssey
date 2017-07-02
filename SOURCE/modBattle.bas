Attribute VB_Name = "modBattle"
Option Explicit

Public Sub Hunt(Creature1 As String, Creature2 As String, Creature3 As String, Creature4 As String, MultiCreature1 As String, MultiCreature2 As String, MultiCreature3 As String, MultiCreature4 As String, Boss As String)

Dim RandCreature As Integer

RandCreature = Rand(1, 100)

If RandCreature <= 15 Then
 Call SingleCreature(Creature1, True)
End If


If RandCreature <= 30 Then
 If RandCreature > 15 Then
  Call SingleCreature(Creature2, True)
 End If
End If

If RandCreature <= 45 Then
 If RandCreature > 30 Then
  Call SingleCreature(Creature3, True)
 End If
End If


If RandCreature <= 60 Then
 If RandCreature > 45 Then
  Call SingleCreature(Creature4, True)
 End If
End If




'for multi battle
If RandCreature <= 95 Then
 If RandCreature > 60 Then
Call MultiCreatureCheck(MultiCreature1, MultiCreature2, MultiCreature3, MultiCreature4)
 End If
End If

'for boss

 If RandCreature > 95 Then
Call BossCreatureDatabase(Boss, True)
 End If


Exit Sub








End Sub



Public Sub InBattle()

Compass = False


End Sub

Public Sub DoneBattle()

Compass = True

End Sub

Public Sub WinBattle(ExperienceGiven As Integer, GoldGiven As Integer)

Call AddText("Experience Gained: " & ExperienceGiven, GREEN, 9, False, False, False)
Call AddText("Gold Gained: " & GoldGiven, GREEN, 9, False, False, False)
Experience = Experience + ExperienceGiven
Gold = Gold + GoldGiven





End Sub

Public Function PlayerDamage(EnemyDefence As Integer) As Integer

PlayerDamage = (Strength + Rand(1, 5)) - EnemyDefence

If PlayerDamage < 0 Then
 PlayerDamage = 0
  End If
  
PlayerDamageTXT = PlayerDamage

End Function

Public Function EnemyDamage(EnemyStrength As Integer) As Integer

EnemyDamage = (EnemyStrength + Rand(1, 5)) - (Defence / 2)

If EnemyDamage < 0 Then
 EnemyDamage = 0
  End If

EnemyDamageTXT = EnemyDamage

 

End Function



Public Sub ResetCreature()

       CreatureName = vbNullString
        CreatureHealth = 0
         CreatureStrength = 0
          CreatureDefence = 0
           CreatureExperience = 0
            CreatureGold = 0
             CreaturePosion = False
             Call DoneBattle
        Exit Sub

End Sub


Public Sub FullBattle(CanFlee As Boolean)


'----------------------------------------------------------------------------------------------------

'Initial Battle Text
   Call AddText("======================", GREEN, 9, False, True, False)
    Call AddText("===== -A " & CreatureName & " Appears- =====", GREEN, 9, False, True, False)
     Call AddText("======================", GREEN, 9, False, True, False)
      Call AddText("-HP: " & CreatureHealth & "-", GREEN, 9, False, True, False)
       Call AddText("", BLUE, 9, False, False, False)
        Call AddText("-Battle Commands-", GREEN, 9, False, True, True)
         Call AddText("", BLUE, 9, False, False, False)
          Call AddText("Attack // Flee", GREEN, 9, False, True, False)
           Call AddText("", BLUE, 9, False, False, False)
           
'Set compass to false so they cant leave location
  Call InBattle
         
'Actual Battle--------
CreatureLoop:
    Do While Compass = "False"
     
     If Health <= 0 Then
      Call ManualDeathCheck
       End If
       
     'If the creature dies-------------------
     If CreatureHealth <= 0 Then
      Call AddText("-You have slain the " & CreatureName & " !", GREEN, 9, False, False, True)
       Call AddText("", BLUE, 9, False, False, False)
        Call WinBattle(CreatureExperience, CreatureGold)
         Call NullPText
          Call ResetCreature
           Call CheckLevelUp
            Call SaveGame
             Call CheckEvents
              Exit Sub
           End If
      '--------------------------------------
      
      
      'check for attack
      If PText = "attack" Then
        PlayerMiss = Rand(1, 100)
        CreatureMiss = Rand(1, 100)
        
'Calculate Enemy Attack-------------------
             
        'if creature hit
           If CreatureMiss >= 10 Then
            If Health > 0 Then
           Health = Health - EnemyDamage(CreatureStrength)
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & CreatureName & " attacks you for " & EnemyDamageTXT & " damage!", RED, 9, False, True, True)
              Call AddText("-You have " & Health & " health left!", RED, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
               
               
               
               '-----Posion check
                If CreaturePosion = True Then
                 If Health > 0 Then
                 
                 CreaturePosionHit = Rand(1, 10)
                  
                  If CreaturePosionHit >= 7 Then
                   
                  Health = Health - EnemyPosion(CreatureStrength)
                   Call AddText("-You Have Been Effected By The " & CreatureName & "'s Posion-", DARKGREY, 9, False, True, True)
                    Call AddText("-You Take " & EnemyDamageTXT & " Damage From The Posion-", DARKGREY, 9, False, True, True)
                     Call AddText("-You have " & Health & " health left!", DARKGREY, 9, False, True, True)
                      Call AddText("", BLUE, 9, False, False, False)
                     End If
                    End If
                   End If
                '-----------------
                
                End If
               End If
               
        'if creature miss
         If CreatureMiss < 10 Then
          If Health > 0 Then
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("-The " & CreatureName & " attacks but missed you!", RED, 9, False, True, True)
             Call AddText("", BLUE, 9, False, False, False)
              End If
             End If
'-----------------------------------------
         
         
         
'Calculate Player Attack------------------



              
           
           
       'If Player Hits
        If PlayerMiss >= (10 - (Speed / 10)) Then
         If Health > 0 Then
         CreatureHealth = CreatureHealth - PlayerDamage(CreatureDefence)
         

            Call AddText("-You attack the " & CreatureName & " and deal " & PlayerDamageTXT & " damage!-", GREEN, 9, False, True, True)
              Call AddText("-The " & CreatureName & " has " & CreatureHealth & " health left!", GREEN, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
                Call NullPText
                   GoTo CreatureLoop
                    End If
                     End If
                     
                'if player miss
         If PlayerMiss < (10 - (Speed / 10)) Then
          If Health > 0 Then
           Call AddText("-You attack the " & CreatureName & " but miss!", GREEN, 9, False, True, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call NullPText
              GoTo CreatureLoop
               End If
              End If
             End If
             
        'if they want to flee
          If PText = "flee" Then
           If CanFlee = True Then
           Call AddText("-You run from the enemy!", GREEN, 9, False, False, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call NullPText
              Call DoneBattle
               Call SaveGame
                Call CheckEvents
              Exit Sub
             End If
            End If
            
        If PText = "flee" Then
         If CanFlee = False Then
          Call AddText("-You Cannot Run From This Enemy!!!-", GREEN, 9, False, False, True)
           Call AddText("", BLUE, 9, False, False, False)
            Call NullPText
             GoTo CreatureLoop
           End If
          End If
    

     If Health <= 0 Then
      Call ManualDeathCheck
       End If
     
 
      DoEvents
    Loop

End Sub


Public Sub MultiCreatureCheck(MultiCreat1 As String, MultiCreat2 As String, MultiCreat3 As String, MultiCreat4 As String)

Dim CreaturePicker As Integer

CreaturePicker = Rand(1, 4)

MultiCreat1 = LCase(MultiCreat1)
MultiCreat2 = LCase(MultiCreat2)
MultiCreat3 = LCase(MultiCreat3)
MultiCreat4 = LCase(MultiCreat4)

'1st Creature
If CreaturePicker = 1 Then
 Call MultiCreature1Database(MultiCreat1)
End If

If CreaturePicker = 2 Then
 Call MultiCreature1Database(MultiCreat2)
End If

If CreaturePicker = 3 Then
 Call MultiCreature1Database(MultiCreat3)
End If
 
If CreaturePicker = 4 Then
 Call MultiCreature1Database(MultiCreat4)
End If

'Reroll creature picker
CreaturePicker = Rand(1, 4)


'2nd Creature
If CreaturePicker = 1 Then
 Call MultiCreature2Database(MultiCreat1)
End If

If CreaturePicker = 2 Then
 Call MultiCreature2Database(MultiCreat2)
End If

If CreaturePicker = 3 Then
 Call MultiCreature2Database(MultiCreat3)
End If
 
If CreaturePicker = 4 Then
 Call MultiCreature2Database(MultiCreat4)
End If



If MultiCreature1Health > 0 Then
 If MultiCreature2Health > 0 Then
  Call TwoCreatureBattle
   End If
    End If
   



End Sub

Public Sub TwoCreatureBattle()

Dim Creature1DeadCheck As Boolean
Dim Creature2DeadCheck As Boolean



   Call AddText("======================", GREEN, 9, False, True, False)
    Call AddText("===== -A " & "(1)" & MultiCreature1Name & " And " & "(2)" & MultiCreature2Name & " Appears- =====", GREEN, 9, False, True, False)
     Call AddText("======================", GREEN, 9, False, True, False)
      Call AddText("(1)" & MultiCreature1Name & " -HP: " & MultiCreature1Health & " -", GREEN, 9, False, True, False)
       Call AddText("", BLUE, 9, False, False, False)
        Call AddText("(2)" & MultiCreature2Name & " -HP: " & MultiCreature2Health & " -", GREEN, 9, False, True, False)
         Call AddText("", BLUE, 9, False, False, False)
          Call AddText("-Battle Commands-", GREEN, 9, False, True, True)
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("Attack (Creature Number) // Flee", GREEN, 9, False, True, False)
             Call AddText("", BLUE, 9, False, False, False)

       'In Battle
        Call InBattle
         Creature1DeadCheck = False
          Creature2DeadCheck = False



CreatureLoop:
    Do While Compass = "False"
    
          If Health <= 0 Then
      Call ManualDeathCheck
       End If
     
     'If the 1st creature dies-------------------
     If MultiCreature1Health <= 0 Then
      If Creature1DeadCheck = False Then
       Call AddText("-You have slain the " & "(1)" & MultiCreature1Name & " !", GREEN, 9, False, False, True)
        Call AddText("", BLUE, 9, False, False, False)
         Call AddText("-You have gained " & MultiCreature1Experience & " Experience and " & MultiCreature1Gold & " Gold!", GREEN, 9, False, False, True)
          Call NullPText
           Call WinBattle(MultiCreature1Experience, MultiCreature1Gold)
    
            
           If Creature2DeadCheck = False Then
           Call AddText("", BLUE, 9, False, False, False)
           Call AddText("-Battle Commands-", GREEN, 9, False, True, True)
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("Attack (Creature Number) // Flee", GREEN, 9, False, True, False)
             Call AddText("", BLUE, 9, False, False, False)
              End If
              
          Call ResetMultiCreature1
           Creature1DeadCheck = True
           End If
          End If
      '--------------------------------------
      
      
      
      'If the 2nd creature dies-------------------
     If MultiCreature2Health <= 0 Then
      If Creature2DeadCheck = False Then
       Call AddText("-You have slain the " & "(2)" & MultiCreature2Name & " !", GREEN, 9, False, False, True)
        Call AddText("", BLUE, 9, False, False, False)
         Call AddText("-You have gained " & MultiCreature2Experience & " Experience and " & MultiCreature2Gold & " Gold!", GREEN, 9, False, False, True)
          Call NullPText
           Call WinBattle(MultiCreature2Experience, MultiCreature2Gold)
            
            If Creature1DeadCheck = False Then
            Call AddText("", BLUE, 9, False, False, False)
           Call AddText("-Battle Commands-", GREEN, 9, False, True, True)
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("Attack (Creature Number) // Flee", GREEN, 9, False, True, False)
             Call AddText("", BLUE, 9, False, False, False)
              End If
              
          Call ResetMultiCreature2
           Creature2DeadCheck = True
          End If
         End If
      '--------------------------------------
      
      
      
      'Check to see if player wins
        If Creature1DeadCheck = True Then
         If Creature2DeadCheck = True Then
          Call AddText("-You Have Defeated The Enemies!", GREEN, 9, False, False, True)
           Call NullPText
            Call DoneBattle
             Call CheckLevelUp
              Call SaveGame
               Call CheckEvents
                Exit Sub
             End If
              End If

          
          
          
          
          
          
      
      
      'check for attack on creature 1
      If PText = "attack 1" Then
        PlayerMiss = Rand(1, 100)
        CreatureMiss = Rand(1, 100)
        
      If Creature1DeadCheck = True Then
       Call NullPText
        GoTo CreatureLoop
      End If
        
'Calculate Enemy Attack-------------------
             
        'if creature 1 hit
           If CreatureMiss >= 10 Then
            If Health > 0 Then
            If Creature1DeadCheck = False Then
           Health = Health - EnemyDamage(MultiCreature1Strength)
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & MultiCreature1Name & " attacks you for " & EnemyDamageTXT & " damage!", RED, 9, False, True, True)
              Call AddText("-You have " & Health & " health left!", RED, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
               
               '-----Posion check
                If MultiCreature1Posion = True Then
                 If Health > 0 Then
                 MultiCreature1PosionHit = Rand(1, 10)
                  
                  If MultiCreature1PosionHit >= 7 Then
                   
                  Health = Health - EnemyPosion(MultiCreature1Strength)
                   Call AddText("-You Have Been Effected By The " & MultiCreature1Name & "'s Posion-", DARKGREY, 9, False, True, True)
                    Call AddText("-You Take " & EnemyDamageTXT & " Damage From The Posion-", DARKGREY, 9, False, True, True)
                     Call AddText("-You have " & Health & " health left!", DARKGREY, 9, False, True, True)
                      Call AddText("", BLUE, 9, False, False, False)
                     End If
                    End If
                   End If
                '-----------------
               
              End If
             End If
            End If
            
        'if creature 1 miss
         If CreatureMiss < 10 Then
          If Creature1DeadCheck = False Then
           If Health > 0 Then
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("-The " & MultiCreature1Name & " attacks but missed you!", RED, 9, False, True, True)
             Call AddText("", BLUE, 9, False, False, False)
              End If
               End If
                End If
                
              '-----
        CreatureMiss = Rand(1, 100)
              '-----
              
            'if creature 2 hit
           If CreatureMiss >= 10 Then
            If Creature2DeadCheck = False Then
             If Health > 0 Then
           Health = Health - EnemyDamage(MultiCreature2Strength)
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & MultiCreature2Name & " attacks you for " & EnemyDamageTXT & " damage!", RED, 9, False, True, True)
              Call AddText("-You have " & Health & " health left!", RED, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
               
               '-----Posion check
                If MultiCreature2Posion = True Then
                 If Health > 0 Then
                
                 MultiCreature2PosionHit = Rand(1, 10)
                  
                  If MultiCreature2PosionHit >= 7 Then
                   
                  Health = Health - EnemyPosion(MultiCreature2Strength)
                   Call AddText("-You Have Been Effected By The " & MultiCreature2Name & "'s Posion-", DARKGREY, 9, False, True, True)
                    Call AddText("-You Take " & EnemyDamageTXT & " Damage From The Posion-", DARKGREY, 9, False, True, True)
                     Call AddText("-You have " & Health & " health left!", DARKGREY, 9, False, True, True)
                      Call AddText("", BLUE, 9, False, False, False)
                     End If
                    End If
                   End If
                '-----------------
               
                End If
               End If
              End If
               
        'if creature 2 miss
         If CreatureMiss < 10 Then
          If Health > 0 Then
           If Creature2DeadCheck = False Then
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & MultiCreature2Name & " attacks but missed you!", RED, 9, False, True, True)
              Call AddText("", BLUE, 9, False, False, False)
               End If
                End If
                 End If
'-----------------------------------------
         
         
         
'Calculate Player Attack------------------





       'If Player Hits
        
        If PlayerMiss >= (10 - (Speed / 10)) Then
         If Health > 0 Then
         MultiCreature1Health = MultiCreature1Health - PlayerDamage(MultiCreature1Defence)
         

            Call AddText("-You attack the " & MultiCreature1Name & " and deal " & PlayerDamageTXT & " damage!-", GREEN, 9, False, True, True)
              Call AddText("-The " & MultiCreature1Name & " has " & MultiCreature1Health & " health left!", GREEN, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
                Call NullPText
                   GoTo CreatureLoop
                    End If
                     End If
    
                'if player miss
         If PlayerMiss < (10 - (Speed / 10)) Then
          If Health > 0 Then
           Call AddText("-You attack the " & MultiCreature1Name & " but miss!", GREEN, 9, False, True, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call NullPText
              GoTo CreatureLoop
               End If
              End If
             End If
    '-------------------------------------------------
    '0000000000000000000000000000000000000000000000000
    '-------------------------------------------------
    
    
    
      'check for attack on creature 2
      If PText = "attack 2" Then
        PlayerMiss = Rand(1, 100)
        CreatureMiss = Rand(1, 100)
        
      'make sure the creature is alive
      If Creature2DeadCheck = True Then
       Call NullPText
        GoTo CreatureLoop
      End If
'Calculate Enemy Attack-------------------
             
        'if creature 1 hit
           If CreatureMiss >= 10 Then
            If Health > 0 Then
            If Creature1DeadCheck = False Then
           Health = Health - EnemyDamage(MultiCreature1Strength)
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & MultiCreature1Name & " attacks you for " & EnemyDamageTXT & " damage!", RED, 9, False, True, True)
              Call AddText("-You have " & Health & " health left!", RED, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
               
               '-----Posion check
                If MultiCreature1Posion = True Then
                 If Health > 0 Then
                 MultiCreature1PosionHit = Rand(1, 10)
                  
                  If MultiCreature1PosionHit >= 7 Then
                   
                  Health = Health - EnemyPosion(MultiCreature1Strength)
                   Call AddText("-You Have Been Effected By The " & MultiCreature1Name & "'s Posion-", DARKGREY, 9, False, True, True)
                    Call AddText("-You Take " & EnemyDamageTXT & " Damage From The Posion-", DARKGREY, 9, False, True, True)
                     Call AddText("-You have " & Health & " health left!", DARKGREY, 9, False, True, True)
                      Call AddText("", BLUE, 9, False, False, False)
                     End If
                    End If
                   End If
                '-----------------
               
                End If
               End If
              End If
               
        'if creature 1 miss
         If CreatureMiss < 10 Then
          If Health > 0 Then
          If Creature1DeadCheck = False Then
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("-The " & MultiCreature1Name & " attacks but missed you!", RED, 9, False, True, True)
             Call AddText("", BLUE, 9, False, False, False)
              End If
               End If
                End If
               
              '-----
        CreatureMiss = Rand(1, 100)
              '-----
              
            'if creature 2 hit
           If CreatureMiss >= 10 Then
            If Health > 0 Then
            If Creature2DeadCheck = False Then
           Health = Health - EnemyDamage(MultiCreature2Strength)
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & MultiCreature2Name & " attacks you for " & EnemyDamageTXT & " damage!", RED, 9, False, True, True)
              Call AddText("-You have " & Health & " health left!", RED, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
               
               '-----Posion check
                If MultiCreature2Posion = True Then
                 If Health > 0 Then
                 
                 MultiCreature2PosionHit = Rand(1, 10)
                  
                  If MultiCreature2PosionHit >= 7 Then
                   
                  Health = Health - EnemyPosion(MultiCreature2Strength)
                   Call AddText("-You Have Been Effected By The " & MultiCreature2Name & "'s Posion-", DARKGREY, 9, False, True, True)
                    Call AddText("-You Take " & EnemyDamageTXT & " Damage From The Posion-", DARKGREY, 9, False, True, True)
                     Call AddText("-You have " & Health & " health left!", DARKGREY, 9, False, True, True)
                      Call AddText("", BLUE, 9, False, False, False)
                     End If
                    End If
                   End If
                '-----------------
               
                End If
               End If
              End If
               
        'if creature 2 miss
         If CreatureMiss < 10 Then
          If Health > 0 Then
          If Creature2DeadCheck = False Then
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("-The " & MultiCreature2Name & " attacks but missed you!", RED, 9, False, True, True)
             Call AddText("", BLUE, 9, False, False, False)
              End If
               End If
                End If
                
'-----------------------------------------
         
         
         
'Calculate Player Attack------------------


       'If Player Hits
        
        If PlayerMiss >= (10 - (Speed / 10)) Then
         If Health > 0 Then
         MultiCreature2Health = MultiCreature2Health - PlayerDamage(MultiCreature2Defence)
         

            Call AddText("-You attack the " & MultiCreature2Name & " and deal " & PlayerDamageTXT & " damage!-", GREEN, 9, False, True, True)
              Call AddText("-The " & MultiCreature2Name & " has " & MultiCreature2Health & " health left!", GREEN, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
                Call NullPText
                   GoTo CreatureLoop
                    End If
                     End If
                     
                'if player miss
         If PlayerMiss < (10 - (Speed / 10)) Then
          If Health > 0 Then
           Call AddText("-You attack the " & MultiCreature2Name & " but miss!", GREEN, 9, False, True, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call NullPText
              GoTo CreatureLoop
               End If
              End If
             End If
'------------------------------------------------------------------

        'if they want to flee
          If PText = "flee" Then
           Call AddText("-You run from the enemies!", GREEN, 9, False, False, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call DoneBattle
              Call ResetMultiCreature1
               Call ResetMultiCreature2
                Call NullPText
               Call CheckEvents
              Exit Sub
             End If
     
     
 
      DoEvents
    Loop

     If Health <= 0 Then
      Call ManualDeathCheck
       End If

End Sub



Public Sub ResetMultiCreature1()

MultiCreature1Name = ""
MultiCreature1Health = 0
MultiCreature1Strength = 0
MultiCreature1Defence = 0
MultiCreature1Experience = 0
MultiCreature1Gold = 0
MultiCreature1Posion = False

End Sub

Public Sub ResetMultiCreature2()

MultiCreature2Name = ""
MultiCreature2Health = 0
MultiCreature2Strength = 0
MultiCreature2Defence = 0
MultiCreature2Experience = 0
MultiCreature2Gold = 0
MultiCreature2Posion = False

End Sub



Public Sub BossBattle(CanFlee As Boolean)


'----------------------------------------------------------------------------------------------------

'Initial Battle Text
   Call AddText("======================", GREEN, 9, False, True, False)
    Call AddText("===== -A " & CreatureName & " Appears- =====", GREEN, 9, False, True, False)
     Call AddText("======================", GREEN, 9, False, True, False)
      Call AddText("-HP: " & CreatureHealth & "-", GREEN, 9, False, True, False)
       Call AddText("", BLUE, 9, False, False, False)
        Call AddText("-Battle Commands-", GREEN, 9, False, True, True)
         Call AddText("", BLUE, 9, False, False, False)
          Call AddText("Attack // Flee", GREEN, 9, False, True, False)
           Call AddText("", BLUE, 9, False, False, False)
           
'Set compass to false so they cant leave location
  Call InBattle
         
'Actual Battle--------
CreatureLoop:
    Do While Compass = "False"
     
     If Health <= 0 Then
      Call ManualDeathCheck
       End If
     
     'If the creature dies-------------------
     If CreatureHealth <= 0 Then
      Call AddText("-You have slain the " & CreatureName & " !", GREEN, 9, False, False, True)
       Call AddText("", BLUE, 9, False, False, False)
        Call WinBattle(CreatureExperience, CreatureGold)
         Call NullPText
          Call ResetCreature
           Call CheckLevelUp
            Call SaveGame
             Call CheckEvents
              Exit Sub
           End If
      '--------------------------------------
      
      
      'check for attack
      If PText = "attack" Then
        PlayerMiss = Rand(1, 100)
        CreatureMiss = Rand(1, 100)
        
'Calculate Enemy Attack-------------------
             
        'if creature hit
           If CreatureMiss >= 10 Then
            If Health > 0 Then
           Health = Health - EnemyDamage(CreatureStrength)
            Call AddText("", BLUE, 9, False, False, False)
             Call AddText("-The " & CreatureName & " attacks you for " & EnemyDamageTXT & " damage!", RED, 9, False, True, True)
              Call AddText("-You have " & Health & " health left!", RED, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
               
               
               '-----Posion check
                If CreaturePosion = True Then
                 If Health > 0 Then
                
                 CreaturePosionHit = Rand(1, 10)
                  
                  If CreaturePosionHit >= 7 Then
                   
                  Health = Health - EnemyPosion(CreatureStrength)
                   Call AddText("-You Have Been Effected By The " & CreatureName & "'s Posion-", DARKGREY, 9, False, True, True)
                    Call AddText("-You Take " & EnemyDamageTXT & " Damage From The Posion-", DARKGREY, 9, False, True, True)
                     Call AddText("-You have " & Health & " health left!", DARKGREY, 9, False, True, True)
                      Call AddText("", BLUE, 9, False, False, False)
                     End If
                    End If
                   End If
                '-----------------
                  
               
                End If
               End If
               
        'if creature miss
         If CreatureMiss < 10 Then
          If Health > 0 Then
           Call AddText("", BLUE, 9, False, False, False)
            Call AddText("-The " & CreatureName & " attacks but missed you!", RED, 9, False, True, True)
             Call AddText("", BLUE, 9, False, False, False)
              End If
             End If
             
'-----------------------------------------
         
         
         
'Calculate Player Attack------------------



              
           
           
       'If Player Hits
        If PlayerMiss >= (10 - (Speed / 10)) Then
         If Health > 0 Then
         CreatureHealth = CreatureHealth - PlayerDamage(CreatureDefence)
         

            Call AddText("-You attack the " & CreatureName & " and deal " & PlayerDamageTXT & " damage!-", GREEN, 9, False, True, True)
              Call AddText("-The " & CreatureName & " has " & CreatureHealth & " health left!", GREEN, 9, False, True, True)
               Call AddText("", BLUE, 9, False, False, False)
                Call NullPText
                   GoTo CreatureLoop
                    End If
                     End If
                     
                'if player miss
         If PlayerMiss < (10 - (Speed / 10)) Then
          If Health > 0 Then
           Call AddText("-You attack the " & CreatureName & " but miss!", GREEN, 9, False, True, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call NullPText
              GoTo CreatureLoop
               End If
              End If
             End If
             
        'if they want to flee
          If PText = "flee" Then
           If CanFlee = True Then
           Call AddText("-You run from the enemy!", GREEN, 9, False, False, True)
            Call AddText("", BLUE, 9, False, False, False)
             Call NullPText
              Call DoneBattle
               Call SaveGame
                Call CheckEvents
              Exit Sub
             End If
            End If
            
        If PText = "flee" Then
         If CanFlee = False Then
          Call AddText("-You Cannot Run From This Enemy!!!-", GREEN, 9, False, False, True)
           Call AddText("", BLUE, 9, False, False, False)
            Call NullPText
             GoTo CreatureLoop
           End If
          End If
    

     If Health <= 0 Then
      Call ManualDeathCheck
       End If
     
 
      DoEvents
    Loop

End Sub

Public Sub CompletedQuest(ExperienceGiven As Integer, GoldGiven As Integer)

Call AddText("//You Have Completed A Quest\\", GREEN, 9, False, False, False)
 Call AddText("You Are Rewarded With " & GoldGiven & " Gold!", GREEN, 9, False, False, False)
  Call AddText("You Are Rewarded With " & ExperienceGiven & " Experience!", GREEN, 9, False, False, False)
   Experience = Experience + ExperienceGiven
    Gold = Gold + GoldGiven
   Call CheckLevelUp
  Call SaveGame

End Sub


Public Function EnemyPosion(EnemyStrength As Integer) As Integer

EnemyPosion = ((EnemyStrength + Rand(1, 5)) / 4)

If EnemyPosion < 0 Then
 EnemyPosion = 0
  End If

EnemyDamageTXT = EnemyPosion


End Function
