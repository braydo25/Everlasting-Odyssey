Attribute VB_Name = "modItemHandle"
Option Explicit


Public Sub InputInvOnNew()

Dim i As Integer

For i = 1 To 25

 Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem" & i, "")
  Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem" & i & "Amount", "")
 
  Next i
  

End Sub

Public Sub LoadInventory()

 
 InventoryItem1 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem1")
  InventoryItem1Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem1Amount")

 InventoryItem2 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem2")
  InventoryItem2Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem2Amount")

 InventoryItem3 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem3")
  InventoryItem3Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem3Amount")

 InventoryItem4 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem4")
  InventoryItem4Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem4Amount")

 InventoryItem5 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem5")
  InventoryItem5Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem5Amount")

 InventoryItem6 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem6")
  InventoryItem6Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem6Amount")

 InventoryItem7 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem7")
  InventoryItem7Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem7Amount")

 InventoryItem8 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem8")
  InventoryItem8Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem8Amount")

 InventoryItem9 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem9")
  InventoryItem9Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem9Amount")

 InventoryItem10 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem10")
  InventoryItem10Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem10Amount")

 InventoryItem11 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem11")
  InventoryItem11Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem11Amount")

 InventoryItem12 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem12")
  InventoryItem12Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem12Amount")

 InventoryItem13 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem13")
  InventoryItem13Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem13Amount")

 InventoryItem14 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem14")
  InventoryItem14Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem14Amount")

 InventoryItem15 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem15")
  InventoryItem15Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem15Amount")

 InventoryItem16 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem16")
  InventoryItem16Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem16Amount")

 InventoryItem17 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem17")
  InventoryItem17Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem17Amount")

 InventoryItem18 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem18")
  InventoryItem18Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem18Amount")

 InventoryItem19 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem19")
  InventoryItem19Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem19Amount")

 InventoryItem20 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem20")
  InventoryItem20Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem20Amount")

 InventoryItem21 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem21")
  InventoryItem21Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem21Amount")

 InventoryItem22 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem22")
  InventoryItem22Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem22Amount")

 InventoryItem23 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem23")
  InventoryItem23Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem23Amount")

 InventoryItem24 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem24")
  InventoryItem24Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem24Amount")

 InventoryItem25 = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem25")
  InventoryItem25Amount = GetVar(App.Path & "\Characters\" & CName & ".EoC", "Character", "InventoryItem25Amount")


End Sub

Public Sub DisplayInventory()


Call AddText(CName & "'s Inventory", GREEN, 10, False, True, False)

  
Call AddText("(1) " & InventoryItem1 & " || Quantity: " & InventoryItem1Amount, DARKGREY, 8, False, False, False)

Call AddText("(2) " & InventoryItem2 & " || Quantity: " & InventoryItem2Amount, DARKGREY, 8, False, False, False)

Call AddText("(3) " & InventoryItem3 & " || Quantity: " & InventoryItem3Amount, DARKGREY, 8, False, False, False)

Call AddText("(4) " & InventoryItem4 & " || Quantity: " & InventoryItem4Amount, DARKGREY, 8, False, False, False)

Call AddText("(5) " & InventoryItem5 & " || Quantity: " & InventoryItem5Amount, DARKGREY, 8, False, False, False)

Call AddText("(6) " & InventoryItem6 & " || Quantity: " & InventoryItem6Amount, DARKGREY, 8, False, False, False)

Call AddText("(7) " & InventoryItem7 & " || Quantity: " & InventoryItem7Amount, DARKGREY, 8, False, False, False)

Call AddText("(8) " & InventoryItem8 & " || Quantity: " & InventoryItem8Amount, DARKGREY, 8, False, False, False)

Call AddText("(9) " & InventoryItem9 & " || Quantity: " & InventoryItem9Amount, DARKGREY, 8, False, False, False)

Call AddText("(10) " & InventoryItem10 & " || Quantity: " & InventoryItem10Amount, DARKGREY, 8, False, False, False)

Call AddText("(11) " & InventoryItem11 & " || Quantity: " & InventoryItem11Amount, DARKGREY, 8, False, False, False)

Call AddText("(12) " & InventoryItem12 & " || Quantity: " & InventoryItem12Amount, DARKGREY, 8, False, False, False)

Call AddText("(13) " & InventoryItem13 & " || Quantity: " & InventoryItem13Amount, DARKGREY, 8, False, False, False)

Call AddText("(14) " & InventoryItem14 & " || Quantity: " & InventoryItem14Amount, DARKGREY, 8, False, False, False)

Call AddText("(15) " & InventoryItem15 & " || Quantity: " & InventoryItem15Amount, DARKGREY, 8, False, False, False)

Call AddText("(16) " & InventoryItem16 & " || Quantity: " & InventoryItem16Amount, DARKGREY, 8, False, False, False)

Call AddText("(17) " & InventoryItem17 & " || Quantity: " & InventoryItem17Amount, DARKGREY, 8, False, False, False)

Call AddText("(18) " & InventoryItem18 & " || Quantity: " & InventoryItem18Amount, DARKGREY, 8, False, False, False)

Call AddText("(19) " & InventoryItem19 & " || Quantity: " & InventoryItem19Amount, DARKGREY, 8, False, False, False)

Call AddText("(20) " & InventoryItem20 & " || Quantity: " & InventoryItem20Amount, DARKGREY, 8, False, False, False)

Call AddText("(21) " & InventoryItem21 & " || Quantity: " & InventoryItem21Amount, DARKGREY, 8, False, False, False)

Call AddText("(22) " & InventoryItem22 & " || Quantity: " & InventoryItem22Amount, DARKGREY, 8, False, False, False)

Call AddText("(23) " & InventoryItem23 & " || Quantity: " & InventoryItem23Amount, DARKGREY, 8, False, False, False)

Call AddText("(24) " & InventoryItem24 & " || Quantity: " & InventoryItem24Amount, DARKGREY, 8, False, False, False)

Call AddText("(25) " & InventoryItem25 & " || Quantity: " & InventoryItem25Amount, DARKGREY, 8, False, False, False)

End Sub



Public Sub Shop(Item1 As String, Item2 As String, Item3 As String, Item4 As String, Item5 As String, Item6 As String, Item7 As String, Item8 As String, Item9 As String, Item10 As String)

Dim a As Boolean
Dim CalledFLocation As String


Dim Item1Exist As Boolean
Dim Item2Exist As Boolean
Dim Item3Exist As Boolean
Dim Item4Exist As Boolean
Dim Item5Exist As Boolean
Dim Item6Exist As Boolean
Dim Item7Exist As Boolean
Dim Item8Exist As Boolean
Dim Item9Exist As Boolean
Dim Item10Exist As Boolean

Item1Exist = False
Item2Exist = False
Item3Exist = False
Item4Exist = False
Item5Exist = False
Item6Exist = False
Item7Exist = False
Item8Exist = False
Item9Exist = False
Item10Exist = False


'Check If Items Are Used
If Item1 <> "" Then
 Item1Exist = True
End If


If Item2 <> "" Then
 Item2Exist = True
End If


If Item3 <> "" Then
 Item3Exist = True
End If


If Item4 <> "" Then
 Item4Exist = True
End If


If Item5 <> "" Then
 Item5Exist = True
End If


If Item6 <> "" Then
 Item6Exist = True
End If


If Item7 <> "" Then
 Item7Exist = True
End If


If Item8 <> "" Then
 Item8Exist = True
End If


If Item9 <> "" Then
 Item9Exist = True
End If


If Item10 <> "" Then
 Item10Exist = True
End If
'=====================


a = False

CalledFLocation = FLocation

Call AddText("You Are Now In The " & Location & " Shop", BROWN, 9, False, True, True)

ShopStartLoop:
Call AddText("", GREEN, 9, False, False, False)
Call AddText("Shopkeeper: How can I help you?", GREEN, 9, False, False, False)
Call AddText("- - - - - - - - - -", GREEN, 9, False, False, False)
Call AddText("Buy / Sell / Leave", GREEN, 9, False, False, False)
Call AddText("- - - - - - - - - -", GREEN, 9, False, False, False)

Do While a = False

'If they leave the area
If FLocation <> CalledFLocation Then
 Exit Sub
  End If

'if they want to buy

 If PText = "leave" Then
 Call AddText("You Have Left The " & Location & " Shop", BROWN, 9, False, True, True)
  Call CheckEvents
   Exit Sub
  End If


 If PText = "buy" Then
 

 Call NullPText
 
 Call AddText("", GREEN, 9, False, False, False)
 
 Call AddText("--Shopkeeper's Goods--", GREEN, 9, False, False, True)
 
If Item1Exist = True Then
 Call CheckBuyPrice(Item1)
  Call AddText("(1) " & Item1 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
 
If Item2Exist = True Then
 Call CheckBuyPrice(Item2)
  Call AddText("(2) " & Item2 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item3Exist = True Then
 Call CheckBuyPrice(Item3)
  Call AddText("(3) " & Item3 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item4Exist = True Then
 Call CheckBuyPrice(Item4)
  Call AddText("(4) " & Item4 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item5Exist = True Then
 Call CheckBuyPrice(Item5)
  Call AddText("(5) " & Item5 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item6Exist = True Then
 Call CheckBuyPrice(Item6)
  Call AddText("(6) " & Item6 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item7Exist = True Then
 Call CheckBuyPrice(Item7)
  Call AddText("(7) " & Item7 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item8Exist = True Then
 Call CheckBuyPrice(Item8)
  Call AddText("(8) " & Item8 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item9Exist = True Then
 Call CheckBuyPrice(Item9)
  Call AddText("(9) " & Item9 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
    
If Item10Exist = True Then
 Call CheckBuyPrice(Item10)
  Call AddText("(10) " & Item10 & " || Price: " & ItemBuyPrice & " Gold", GREEN, 9, False, False, False)
    End If
 
Call AddText("", GREEN, 9, False, False, False)
Call AddText("Shopkeeper: What Would You Like To Buy?", GREEN, 9, False, False, False)
Call AddText("- - - - - - - - - -", GREEN, 9, False, False, False)
Call AddText("Buy (Item Number) / Description (Item Number) / Stop Buying", GREEN, 9, False, False, False)
Call AddText("- - - - - - - - - -", GREEN, 9, False, False, False)
 
BuyLoop:
 Do While a = False
 
  If PText = "stop buying" Then
   GoTo ShopStartLoop
    End If
 
 
  'buy item 1
  If PText = "buy 1" Then
   If Item1Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item1)
     Call CheckForBlankInv
     
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
   
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item1)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item1, GREEN, 9, False, True, False)
 Call LoadInventory
   
   
  End If
 End If
 
 
 
  'buy item 2
  If PText = "buy 2" Then
   If Item2Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item2)
     Call CheckForBlankInv
   
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
   
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item2)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item2, GREEN, 9, False, True, False)
 Call LoadInventory
   End If
  End If
  
  
  
  'buy item 3
  If PText = "buy 3" Then
   If Item3Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item3)
     Call CheckForBlankInv
  
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
  
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item3)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item3, GREEN, 9, False, True, False)
 Call LoadInventory
   End If
  End If
  
  
  
  'buy item 4
  If PText = "buy 4" Then
   If Item4Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item4)
     Call CheckForBlankInv
     
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
   
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item4)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item4, GREEN, 9, False, True, False)
 Call LoadInventory
   End If
  End If
  
  
  
  'buy item 5
  If PText = "buy 5" Then
   If Item5Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item5)
     Call CheckForBlankInv
  
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
  
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item5)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item5, GREEN, 9, False, True, False)
 Call LoadInventory
  
   End If
  End If
  
  
  
  'buy item 6
  If PText = "buy 6" Then
   If Item6Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item6)
     Call CheckForBlankInv
  
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
  
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item6)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item6, GREEN, 9, False, True, False)
 Call LoadInventory
  
   End If
  End If
  
  
  
  'buy item 7
  If PText = "buy 7" Then
   If Item7Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item7)
     Call CheckForBlankInv
   
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
   
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item7)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item7, GREEN, 9, False, True, False)
 Call LoadInventory
   
   End If
  End If
  
  
  
  'buy item 8
  If PText = "buy 8" Then
   If Item8Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item8)
     Call CheckForBlankInv
  
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
  
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item8)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item8, GREEN, 9, False, True, False)
 Call LoadInventory
  
   End If
  End If
  
  
  
  'buy item 9
  If PText = "buy 9" Then
   If Item9Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item9)
     Call CheckForBlankInv
  
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
  
'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item9)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item9, GREEN, 9, False, True, False)
 Call LoadInventory
  
   End If
  End If
  
  
  
  'buy item 10
  If PText = "buy 10" Then
   If Item10Exist = True Then
   Call NullPText
    Call CheckBuyPrice(Item10)
     Call CheckForBlankInv
  
     If FullInventory = True Then
    Call AddText("Your Inventory Is Full! You Cannot Purchase That Item", RED, 9, False, True, False)
   GoTo BuyLoop
  End If
  
 'check if the have enough money
If Gold < ItemBuyPrice Then
 Call AddText("You Do Not Have Enough Gold To Purchase This Item!", RED, 9, False, True, False)
  GoTo BuyLoop
   End If
   
   
 'take gold
 Gold = Gold - ItemBuyPrice

'give items
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot, Item10)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", OpenInvSlot & "Amount", "1")
 Call AddText("You Have Successfully Purchased The " & Item10, GREEN, 9, False, True, False)
 Call LoadInventory
  
   End If
  End If
 
 
 
 
 
 'item 1 description
  If PText = "description 1" Then
   If Item1Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item1)
      End If
       End If
       
       
       
       
 'item 2 description
  If PText = "description 2" Then
   If Item2Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item2)
      End If
       End If
       
       
       
 'item 3 description
  If PText = "description 3" Then
   If Item3Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item3)
      End If
       End If
       
       
       
 'item 4 description
  If PText = "description 4" Then
   If Item4Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item4)
      End If
       End If
       
       
       
 'item 5 description
  If PText = "description 5" Then
   If Item5Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item5)
      End If
       End If
       
       
    
 
 'item 6 description
  If PText = "description 6" Then
   If Item6Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item6)
      End If
       End If
 
 
 
 'item 7 description
  If PText = "description 7" Then
   If Item7Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item7)
      End If
       End If
 
 
 
 'item 8 description
  If PText = "description 8" Then
   If Item8Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item8)
      End If
       End If
 
 
 
 'item 9 description
  If PText = "description 9" Then
   If Item9Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item9)
      End If
       End If
 
 
 
 'item 10 description
  If PText = "description 10" Then
   If Item10Exist = True Then
    Call NullPText
     Call ItemDescriptionDatabase(Item10)
      End If
       End If
 
 
 DoEvents
 Loop

 End If


'======================================
'======================================
'======================================


 If PText = "sell" Then
  Call NullPText
   
   Call DisplayInventory
    
Call AddText("", GREEN, 9, False, False, False)
 Call AddText("Shopkeeper: What Would You Like To Sell From Your Inventory?", GREEN, 9, False, False, False)
  Call AddText("- - - - - - - - - -", GREEN, 9, False, False, False)
   Call AddText("Sell (Inventory Item Number) / Stop Selling", GREEN, 9, False, False, False)
    Call AddText("- - - - - - - - - -", GREEN, 9, False, False, False)

SellItemLoop:

       Do While a = False
       
        If PText = "stop selling" Then
         GoTo ShopStartLoop
          End If



 If PText = "sell 1" Then
  Call NullPText
   
    If InventoryItem1 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem1)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem1 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem1 = ""
            InventoryItem1Amount = ""
            Call SaveGame
             
             End If
             
             
             
             
             
 If PText = "sell 2" Then
  Call NullPText
   
    If InventoryItem2 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem2)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem2 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem2 = ""
            InventoryItem2Amount = ""
            Call SaveGame
             
             End If
             
             
             
             
 If PText = "sell 3" Then
  Call NullPText
   
    If InventoryItem3 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem3)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem3 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem3 = ""
            InventoryItem3Amount = ""
            Call SaveGame
             
             End If
             
             
             
             
             
 If PText = "sell 4" Then
  Call NullPText
   
    If InventoryItem4 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem4)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem4 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem4 = ""
            InventoryItem4Amount = ""
            Call SaveGame
             
             End If
             
             
             
             
             
 If PText = "sell 5" Then
  Call NullPText
   
    If InventoryItem5 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem5)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem5 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem5 = ""
            InventoryItem5Amount = ""
            Call SaveGame
             
             End If
             
             
             
        

 If PText = "sell 6" Then
  Call NullPText
   
    If InventoryItem6 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem6)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem6 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem6 = ""
            InventoryItem6Amount = ""
            Call SaveGame
             
             End If





 If PText = "sell 7" Then
  Call NullPText
   
    If InventoryItem7 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem7)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem7 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem7 = ""
            InventoryItem7Amount = ""
            Call SaveGame
             
             End If





 If PText = "sell 8" Then
  Call NullPText
   
    If InventoryItem8 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem8)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem8 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem8 = ""
            InventoryItem8Amount = ""
            Call SaveGame
             
             End If





 If PText = "sell 9" Then
  Call NullPText
   
    If InventoryItem9 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem9)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem9 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem9 = ""
            InventoryItem9Amount = ""
            Call SaveGame
             
             End If





 If PText = "sell 10" Then
  Call NullPText
   
    If InventoryItem10 = "" Then
     Call AddText("You Do Not Have An Item To Sell In That Slot!", RED, 9, False, False, False)
      GoTo SellItemLoop
       End If
        
        Call CheckSellPrice(InventoryItem10)
         Gold = Gold + ItemSellPrice
          Call AddText("You Successfully Sold A " & InventoryItem10 & " For " & ItemSellPrice & " Gold!", GREEN, 9, False, False, False)
           InventoryItem10 = ""
            InventoryItem10Amount = ""
            Call SaveGame
             
             End If

















DoEvents
Loop














'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ loop seperator
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/ loop seperator


End If

DoEvents
Loop


End Sub

Public Sub DropItem()


'----
 If PText = "/drop 1" Then
  If InventoryItem1 <> "" Then
  Call AddText("You Have Successfully Dropped All " & InventoryItem1 & " In Inventory Slot 1!", CYAN, 9, False, True, False)

   InventoryItem1 = ""
    InventoryItem1Amount = ""
    Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
       
 '----
 If PText = "/drop 2" Then
  If InventoryItem2 <> "" Then
  Call AddText("You Have Successfully Dropped All " & InventoryItem2 & " In Inventory Slot 2!", CYAN, 9, False, True, False)

  InventoryItem2 = ""
   InventoryItem2Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If

 '----
  If PText = "/drop 3" Then
  If InventoryItem3 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem3 & " In Inventory Slot 3!", CYAN, 9, False, True, False)

  InventoryItem3 = ""
   InventoryItem3Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 4" Then
  If InventoryItem4 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem4 & " In Inventory Slot 4!", CYAN, 9, False, True, False)

  InventoryItem4 = ""
   InventoryItem4Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 5" Then
  If InventoryItem5 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem5 & " In Inventory Slot 5!", CYAN, 9, False, True, False)

  InventoryItem5 = ""
   InventoryItem5Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 6" Then
  If InventoryItem6 <> "" Then
      Call AddText("You Have Successfully Dropped All " & InventoryItem6 & " In Inventory Slot 6!", CYAN, 9, False, True, False)

  InventoryItem6 = ""
   InventoryItem6Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 7" Then
  If InventoryItem7 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem7 & " In Inventory Slot 7!", CYAN, 9, False, True, False)

  InventoryItem7 = ""
   InventoryItem7Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 8" Then
  If InventoryItem8 <> "" Then
        Call AddText("You Have Successfully Dropped All " & InventoryItem8 & " In Inventory Slot 8!", CYAN, 9, False, True, False)

  InventoryItem8 = ""
   InventoryItem8Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
'----
 If PText = "/drop 9" Then
  If InventoryItem9 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem9 & " In Inventory Slot 9!", CYAN, 9, False, True, False)

  InventoryItem9 = ""
   InventoryItem9Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 10" Then
  If InventoryItem10 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem10 & " In Inventory Slot 10!", CYAN, 9, False, True, False)

  InventoryItem10 = ""
   InventoryItem10Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 11" Then
  If InventoryItem11 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem11 & " In Inventory Slot 11!", CYAN, 9, False, True, False)

  InventoryItem11 = ""
   InventoryItem11Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 12" Then
  If InventoryItem12 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem12 & " In Inventory Slot 12!", CYAN, 9, False, True, False)

  InventoryItem12 = ""
   InventoryItem12Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 13" Then
  If InventoryItem13 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem13 & " In Inventory Slot 13!", CYAN, 9, False, True, False)

  InventoryItem13 = ""
   InventoryItem13Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 14" Then
  If InventoryItem14 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem14 & " In Inventory Slot 14!", CYAN, 9, False, True, False)

  InventoryItem14 = ""
   InventoryItem14Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 
 '----
 If PText = "/drop 15" Then
  If InventoryItem15 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem15 & " In Inventory Slot 15!", CYAN, 9, False, True, False)

  InventoryItem15 = ""
   InventoryItem15Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 
 '----
 If PText = "/drop 16" Then
  If InventoryItem16 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem16 & " In Inventory Slot 16!", CYAN, 9, False, True, False)

  InventoryItem16 = ""
   InventoryItem16Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 17" Then
  If InventoryItem17 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem17 & " In Inventory Slot 17!", CYAN, 9, False, True, False)

  InventoryItem17 = ""
   InventoryItem17Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 '----
 If PText = "/drop 18" Then
  If InventoryItem18 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem18 & " In Inventory Slot 18!", CYAN, 9, False, True, False)

  InventoryItem18 = ""
   InventoryItem18Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 
 '----
 If PText = "/drop 19" Then
  If InventoryItem19 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem19 & " In Inventory Slot 19!", CYAN, 9, False, True, False)

  InventoryItem19 = ""
   InventoryItem19Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 
 '----
 If PText = "/drop 20" Then
  If InventoryItem20 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem20 & " In Inventory Slot 20!", CYAN, 9, False, True, False)

  InventoryItem20 = ""
   InventoryItem20Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 
 '----
 If PText = "/drop 21" Then
  If InventoryItem21 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem21 & " In Inventory Slot 21!", CYAN, 9, False, True, False)

  InventoryItem21 = ""
   InventoryItem21Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If
 
 
'----
 If PText = "/drop 22" Then
  If InventoryItem22 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem22 & " In Inventory Slot 22!", CYAN, 9, False, True, False)

  InventoryItem22 = ""
   InventoryItem22Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If


'----
 If PText = "/drop 23" Then
  If InventoryItem23 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem23 & " In Inventory Slot 23!", CYAN, 9, False, True, False)

  InventoryItem23 = ""
   InventoryItem23Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If


'----
 If PText = "/drop 24" Then
  If InventoryItem24 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem24 & " In Inventory Slot 24!", CYAN, 9, False, True, False)

  InventoryItem24 = ""
   InventoryItem24Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If

'----
 If PText = "/drop 25" Then
  If InventoryItem25 <> "" Then
    Call AddText("You Have Successfully Dropped All " & InventoryItem25 & " In Inventory Slot 25!", CYAN, 9, False, True, False)
  
  InventoryItem25 = ""
   InventoryItem25Amount = ""
   Call SaveInventory
    
    Else
     Call AddText("You Do Not Have An Item In That Inventory Slot!", CYAN, 9, False, True, False)
      End If
       End If






End Sub


Public Sub SaveInventory()

Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem1", InventoryItem1)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem1Amount", InventoryItem1Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem2", InventoryItem2)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem2Amount", InventoryItem2Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem3", InventoryItem3)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem3Amount", InventoryItem3Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem4", InventoryItem4)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem4Amount", InventoryItem4Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem5", InventoryItem5)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem5Amount", InventoryItem5Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem6", InventoryItem6)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem6Amount", InventoryItem6Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem7", InventoryItem7)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem7Amount", InventoryItem7Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem8", InventoryItem8)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem8Amount", InventoryItem8Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem9", InventoryItem9)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem9Amount", InventoryItem9Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem10", InventoryItem10)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem10Amount", InventoryItem10Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem11", InventoryItem11)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem11Amount", InventoryItem11Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem12", InventoryItem12)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem12Amount", InventoryItem12Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem13", InventoryItem13)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem13Amount", InventoryItem13Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem14", InventoryItem14)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem14Amount", InventoryItem14Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem15", InventoryItem15)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem15Amount", InventoryItem15Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem16", InventoryItem16)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem16Amount", InventoryItem16Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem17", InventoryItem17)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem17Amount", InventoryItem17Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem18", InventoryItem18)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem18Amount", InventoryItem18Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem19", InventoryItem19)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem19Amount", InventoryItem19Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem20", InventoryItem20)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem20Amount", InventoryItem20Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem21", InventoryItem21)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem21Amount", InventoryItem21Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem22", InventoryItem22)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem22Amount", InventoryItem22Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem23", InventoryItem23)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem23Amount", InventoryItem23Amount)

Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem24", InventoryItem24)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem24Amount", InventoryItem24Amount)


Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem25", InventoryItem25)
Call PutVar((App.Path & "\Characters\" & CName & ".EoC"), "Character", "InventoryItem25Amount", InventoryItem25Amount)


End Sub

Public Sub CheckBuyPrice(Item As String)
       
Dim ItemName As String

ItemName = LCase(Item)

'basic potions
If ItemName = "potion" Then
 ItemBuyPrice = 20
End If

If ItemName = "soul potion" Then
 ItemBuyPrice = 50
End If

If ItemName = "large potion" Then
 ItemBuyPrice = 100
End If

If ItemName = "large soul potion" Then
 ItemBuyPrice = 210
End If


'basic equipment
If ItemName = "copper sword" Then
 ItemBuyPrice = 100
End If

If ItemName = "hide chestplate" Then
 ItemBuyPrice = 70
End If

If ItemName = "hide leggings" Then
 ItemBuyPrice = 65
End If

If ItemName = "fur gloves" Then
 ItemBuyPrice = 60
End If

If ItemName = "fur boots" Then
 ItemBuyPrice = 65
End If














End Sub

Public Sub CheckForBlankInv()

If InventoryItem1 = "" Then
 If InventoryItem1Amount = "" Then
  OpenInvSlot = "InventoryItem1"
   FullInventory = False
   Exit Sub
    End If
     End If
    
If InventoryItem2 = "" Then
 If InventoryItem2Amount = "" Then
  OpenInvSlot = "InventoryItem2"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem3 = "" Then
 If InventoryItem3Amount = "" Then
  OpenInvSlot = "InventoryItem3"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem4 = "" Then
 If InventoryItem4Amount = "" Then
  OpenInvSlot = "InventoryItem4"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem5 = "" Then
 If InventoryItem5Amount = "" Then
  OpenInvSlot = "InventoryItem5"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem6 = "" Then
 If InventoryItem6Amount = "" Then
  OpenInvSlot = "InventoryItem6"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem7 = "" Then
 If InventoryItem7Amount = "" Then
  OpenInvSlot = "InventoryItem7"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem8 = "" Then
 If InventoryItem8Amount = "" Then
  OpenInvSlot = "InventoryItem8"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem9 = "" Then
 If InventoryItem9Amount = "" Then
  OpenInvSlot = "InventoryItem9"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem10 = "" Then
 If InventoryItem10Amount = "" Then
  OpenInvSlot = "InventoryItem10"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem11 = "" Then
 If InventoryItem11Amount = "" Then
  OpenInvSlot = "InventoryItem11"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem12 = "" Then
 If InventoryItem12Amount = "" Then
  OpenInvSlot = "InventoryItem12"
   FullInventory = False
   Exit Sub
    End If
     End If

If InventoryItem13 = "" Then
 If InventoryItem13Amount = "" Then
  OpenInvSlot = "InventoryItem13"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem14 = "" Then
 If InventoryItem14Amount = "" Then
  OpenInvSlot = "InventoryItem14"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem15 = "" Then
 If InventoryItem15Amount = "" Then
  OpenInvSlot = "InventoryItem15"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem16 = "" Then
 If InventoryItem16Amount = "" Then
  OpenInvSlot = "InventoryItem16"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem17 = "" Then
 If InventoryItem17Amount = "" Then
  OpenInvSlot = "InventoryItem17"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem18 = "" Then
 If InventoryItem18Amount = "" Then
  OpenInvSlot = "InventoryItem18"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem19 = "" Then
 If InventoryItem19Amount = "" Then
  OpenInvSlot = "InventoryItem19"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem20 = "" Then
 If InventoryItem20Amount = "" Then
  OpenInvSlot = "InventoryItem20"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem21 = "" Then
 If InventoryItem21Amount = "" Then
  OpenInvSlot = "InventoryItem21"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem22 = "" Then
 If InventoryItem22Amount = "" Then
  OpenInvSlot = "InventoryItem22"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem23 = "" Then
 If InventoryItem23Amount = "" Then
  OpenInvSlot = "InventoryItem23"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem24 = "" Then
 If InventoryItem24Amount = "" Then
  OpenInvSlot = "InventoryItem24"
   FullInventory = False
   Exit Sub
    End If
     End If
     
If InventoryItem25 = "" Then
 If InventoryItem25Amount = "" Then
  OpenInvSlot = "InventoryItem25"
   FullInventory = False
   Exit Sub
    End If
     End If



FullInventory = True




End Sub

Public Sub UseItem()

ReCheckInput:

If PText = "/use 1" Then

'check if they actually have an item in that slot
 If InventoryItem1 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
     
Call NullPText
  
 Call UsableItemDatabase(InventoryItem1)
  
  If UsableItem = False Then
   Call AddText(InventoryItem1 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem1 & "!", GREEN, 9, False, True, False)
   InventoryItem1Amount = InventoryItem1Amount - 1
     
     'check if they have any left after use
       If InventoryItem1Amount <= 0 Then
        InventoryItem1 = ""
         InventoryItem1Amount = ""
          End If
          
        
        If Compass = True Then
          Call SaveGame
            End If
            
           Call NullPText
End If










If PText = "/use 2" Then

'check if they actually have an item in that slot
 If InventoryItem2 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
 Call NullPText
 Call UsableItemDatabase(InventoryItem2)
  
  If UsableItem = False Then
   Call AddText(InventoryItem2 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem2 & "!", GREEN, 9, False, True, False)
   InventoryItem2Amount = InventoryItem2Amount - 1
     
     'check if they have any left after use
       If InventoryItem2Amount <= 0 Then
        InventoryItem2 = ""
         InventoryItem2Amount = ""
          End If
   
           If Compass = True Then
          Call SaveGame
            End If
   
           Call NullPText
End If











If PText = "/use 3" Then

'check if they actually have an item in that slot
 If InventoryItem3 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
 Call NullPText
 Call UsableItemDatabase(InventoryItem3)
  
  If UsableItem = False Then
   Call AddText(InventoryItem3 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem3 & "!", GREEN, 9, False, True, False)
   InventoryItem3Amount = InventoryItem3Amount - 1
     
     'check if they have any left after use
       If InventoryItem3Amount <= 0 Then
        InventoryItem3 = ""
         InventoryItem3Amount = ""
          End If
          
        If Compass = True Then
          Call SaveGame
            End If


           Call NullPText
End If











If PText = "/use 4" Then

'check if they actually have an item in that slot
 If InventoryItem4 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem4)
  
  If UsableItem = False Then
   Call AddText(InventoryItem4 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem4 & "!", GREEN, 9, False, True, False)
   InventoryItem4Amount = InventoryItem4Amount - 1
     
     'check if they have any left after use
       If InventoryItem4Amount <= 0 Then
        InventoryItem4 = ""
         InventoryItem4Amount = ""
          End If
          
        If Compass = True Then
          Call SaveGame
            End If

           Call NullPText
End If









If PText = "/use 5" Then

'check if they actually have an item in that slot
 If InventoryItem5 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
 Call NullPText
 Call UsableItemDatabase(InventoryItem5)
  
  If UsableItem = False Then
   Call AddText(InventoryItem5 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem5 & "!", GREEN, 9, False, True, False)
   InventoryItem5Amount = InventoryItem5Amount - 1
     
     'check if they have any left after use
       If InventoryItem5Amount <= 0 Then
        InventoryItem5 = ""
         InventoryItem5Amount = ""
          End If
          
        If Compass = True Then
          Call SaveGame
            End If


           Call NullPText
End If












If PText = "/use 6" Then

'check if they actually have an item in that slot
 If InventoryItem6 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
     
Call NullPText
 Call UsableItemDatabase(InventoryItem6)
  
  If UsableItem = False Then
   Call AddText(InventoryItem6 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem6 & "!", GREEN, 9, False, True, False)
   InventoryItem6Amount = InventoryItem6Amount - 1
     
     'check if they have any left after use
       If InventoryItem6Amount <= 0 Then
        InventoryItem6 = ""
         InventoryItem6Amount = ""
          End If
          
        If Compass = True Then
          Call SaveGame
            End If
            
           Call NullPText
End If











If PText = "/use 7" Then

'check if they actually have an item in that slot
 If InventoryItem7 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem7)
  
  If UsableItem = False Then
   Call AddText(InventoryItem7 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem7 & "!", GREEN, 9, False, True, False)
   InventoryItem7Amount = InventoryItem7Amount - 1
     
     'check if they have any left after use
       If InventoryItem7Amount <= 0 Then
        InventoryItem7 = ""
         InventoryItem7Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 8" Then

'check if they actually have an item in that slot
 If InventoryItem8 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem8)
  
  If UsableItem = False Then
   Call AddText(InventoryItem8 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem8 & "!", GREEN, 9, False, True, False)
   InventoryItem8Amount = InventoryItem8Amount - 1
     
     'check if they have any left after use
       If InventoryItem8Amount <= 0 Then
        InventoryItem8 = ""
         InventoryItem8Amount = ""
          End If
          
            If Compass = True Then
          Call SaveGame
            End If
    
           Call NullPText
End If











If PText = "/use 9" Then

'check if they actually have an item in that slot
 If InventoryItem9 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
 Call NullPText
 Call UsableItemDatabase(InventoryItem9)
  
  If UsableItem = False Then
   Call AddText(InventoryItem9 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem9 & "!", GREEN, 9, False, True, False)
   InventoryItem9Amount = InventoryItem9Amount - 1
     
     'check if they have any left after use
       If InventoryItem9Amount <= 0 Then
        InventoryItem9 = ""
         InventoryItem9Amount = ""
          End If
          
                If Compass = True Then
          Call SaveGame
            End If
        
           Call NullPText
End If











If PText = "/use 10" Then

'check if they actually have an item in that slot
 If InventoryItem10 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem10)
  
  If UsableItem = False Then
   Call AddText(InventoryItem10 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem10 & "!", GREEN, 9, False, True, False)
   InventoryItem10Amount = InventoryItem10Amount - 1
     
     'check if they have any left after use
       If InventoryItem10Amount <= 0 Then
        InventoryItem10 = ""
         InventoryItem10Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 11" Then

'check if they actually have an item in that slot
 If InventoryItem11 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem11)
  
  If UsableItem = False Then
   Call AddText(InventoryItem11 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem11 & "!", GREEN, 9, False, True, False)
   InventoryItem11Amount = InventoryItem11Amount - 1
     
     'check if they have any left after use
       If InventoryItem11Amount <= 0 Then
        InventoryItem11 = ""
         InventoryItem11Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 12" Then

'check if they actually have an item in that slot
 If InventoryItem12 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem12)
  
  If UsableItem = False Then
   Call AddText(InventoryItem12 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem12 & "!", GREEN, 9, False, True, False)
   InventoryItem12Amount = InventoryItem12Amount - 1
     
     'check if they have any left after use
       If InventoryItem12Amount <= 0 Then
        InventoryItem12 = ""
         InventoryItem12Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 13" Then

'check if they actually have an item in that slot
 If InventoryItem13 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
     
Call NullPText
 Call UsableItemDatabase(InventoryItem13)
  
  If UsableItem = False Then
   Call AddText(InventoryItem13 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem13 & "!", GREEN, 9, False, True, False)
   InventoryItem13Amount = InventoryItem13Amount - 1
     
     'check if they have any left after use
       If InventoryItem13Amount <= 0 Then
        InventoryItem13 = ""
         InventoryItem13Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 14" Then

'check if they actually have an item in that slot
 If InventoryItem14 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem14)
  
  If UsableItem = False Then
   Call AddText(InventoryItem14 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem14 & "!", GREEN, 9, False, True, False)
   InventoryItem14Amount = InventoryItem14Amount - 1
     
     'check if they have any left after use
       If InventoryItem14Amount <= 0 Then
        InventoryItem14 = ""
         InventoryItem14Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 15" Then

'check if they actually have an item in that slot
 If InventoryItem15 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem15)
  
  If UsableItem = False Then
   Call AddText(InventoryItem15 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem15 & "!", GREEN, 9, False, True, False)
   InventoryItem15Amount = InventoryItem15Amount - 1
     
     'check if they have any left after use
       If InventoryItem15Amount <= 0 Then
        InventoryItem15 = ""
         InventoryItem15Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 16" Then

'check if they actually have an item in that slot
 If InventoryItem16 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem16)
  
  If UsableItem = False Then
   Call AddText(InventoryItem16 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem16 & "!", GREEN, 9, False, True, False)
   InventoryItem16Amount = InventoryItem16Amount - 1
     
     'check if they have any left after use
       If InventoryItem16Amount <= 0 Then
        InventoryItem16 = ""
         InventoryItem16Amount = ""
          End If
          
                 If Compass = True Then
          Call SaveGame
            End If
         
           Call NullPText
End If











If PText = "/use 17" Then

'check if they actually have an item in that slot
 If InventoryItem17 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem17)
  
  If UsableItem = False Then
   Call AddText(InventoryItem17 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem17 & "!", GREEN, 9, False, True, False)
   InventoryItem17Amount = InventoryItem17Amount - 1
     
     'check if they have any left after use
       If InventoryItem17Amount <= 0 Then
        InventoryItem17 = ""
         InventoryItem17Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 18" Then

'check if they actually have an item in that slot
 If InventoryItem18 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem18)
  
  If UsableItem = False Then
   Call AddText(InventoryItem18 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem18 & "!", GREEN, 9, False, True, False)
   InventoryItem18Amount = InventoryItem18Amount - 1
     
     'check if they have any left after use
       If InventoryItem18Amount <= 0 Then
        InventoryItem18 = ""
         InventoryItem18Amount = ""
          End If
          
          Call SaveGame
           Call NullPText
End If











If PText = "/use 19" Then

'check if they actually have an item in that slot
 If InventoryItem19 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem19)
  
  If UsableItem = False Then
   Call AddText(InventoryItem19 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem19 & "!", GREEN, 9, False, True, False)
   InventoryItem19Amount = InventoryItem19Amount - 1
     
     'check if they have any left after use
       If InventoryItem19Amount <= 0 Then
        InventoryItem19 = ""
         InventoryItem19Amount = ""
          End If
          
                 If Compass = True Then
          Call SaveGame
            End If
         
           Call NullPText
End If











If PText = "/use 20" Then

'check if they actually have an item in that slot
 If InventoryItem20 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem20)
  
  If UsableItem = False Then
   Call AddText(InventoryItem20 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem20 & "!", GREEN, 9, False, True, False)
   InventoryItem20Amount = InventoryItem20Amount - 1
     
     'check if they have any left after use
       If InventoryItem20Amount <= 0 Then
        InventoryItem20 = ""
         InventoryItem20Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 21" Then

'check if they actually have an item in that slot
 If InventoryItem21 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem21)
  
  If UsableItem = False Then
   Call AddText(InventoryItem21 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem21 & "!", GREEN, 9, False, True, False)
   InventoryItem21Amount = InventoryItem21Amount - 1
     
     'check if they have any left after use
       If InventoryItem21Amount <= 0 Then
        InventoryItem21 = ""
         InventoryItem21Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 22" Then

'check if they actually have an item in that slot
 If InventoryItem22 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem22)
  
  If UsableItem = False Then
   Call AddText(InventoryItem22 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem22 & "!", GREEN, 9, False, True, False)
   InventoryItem22Amount = InventoryItem22Amount - 1
     
     'check if they have any left after use
       If InventoryItem22Amount <= 0 Then
        InventoryItem22 = ""
         InventoryItem22Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 23" Then

'check if they actually have an item in that slot
 If InventoryItem23 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem23)
  
  If UsableItem = False Then
   Call AddText(InventoryItem23 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem23 & "!", GREEN, 9, False, True, False)
   InventoryItem23Amount = InventoryItem23Amount - 1
     
     'check if they have any left after use
       If InventoryItem23Amount <= 0 Then
        InventoryItem23 = ""
         InventoryItem23Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
          
           Call NullPText
End If











If PText = "/use 24" Then

'check if they actually have an item in that slot
 If InventoryItem24 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem24)
  
  If UsableItem = False Then
   Call AddText(InventoryItem24 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem24 & "!", GREEN, 9, False, True, False)
   InventoryItem24Amount = InventoryItem24Amount - 1
     
     'check if they have any left after use
       If InventoryItem24Amount <= 0 Then
        InventoryItem24 = ""
         InventoryItem24Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
            
           Call NullPText
End If











If PText = "/use 25" Then

'check if they actually have an item in that slot
 If InventoryItem25 = "" Then
  Call AddText("You Do Not Have An Item In This Inventory Slot!", RED, 9, False, True, False)
   Call NullPText
    GoTo ReCheckInput
     End If
       
Call NullPText
 Call UsableItemDatabase(InventoryItem25)
  
  If UsableItem = False Then
   Call AddText(InventoryItem25 & " Is Not A Usable Item!", RED, 9, False, True, False)
    Exit Sub
     End If
  
  Call AddText("You Used A " & InventoryItem25 & "!", GREEN, 9, False, True, False)
   InventoryItem25Amount = InventoryItem25Amount - 1
     
     'check if they have any left after use
       If InventoryItem25Amount <= 0 Then
        InventoryItem25 = ""
         InventoryItem25Amount = ""
          End If
          
                  If Compass = True Then
          Call SaveGame
            End If
            
           Call NullPText
End If




End Sub

Public Sub UsableItemDatabase(Item As String)

Dim pItem As String


pItem = LCase(Item)





If pItem = "potion" Then

If Health + 25 > MaxHealth Then
 Health = MaxHealth
  End If

If Health <> MaxHealth Then
 Health = Health + 25
End If

  UsableItem = True
     
     Exit Sub
End If





If pItem = "soul potion" Then

If MP + 10 > MaxMP Then
 MP = MaxMP
  End If
  
If MP <> MaxMP Then
 MP = MP + 10
  End If

  UsableItem = True
     Exit Sub
End If





If pItem = "large potion" Then

If Health + 150 > MaxHealth Then
 Health = MaxHealth
  End If
  
If Health <> MaxHealth Then
 Health = Health + 150
  End If
  
  UsableItem = True
   Exit Sub
End If





If pItem = "large soul potion" Then

If MP + 50 > MaxMP Then
 MP = MaxMP
  End If
  
If MP <> MaxMP Then
 MP = MP + 50
  End If
 
  UsableItem = True
   Exit Sub
End If







UsableItem = False

End Sub


Public Sub ItemDescriptionDatabase(Item As String)

Dim PickedItem As String

PickedItem = LCase(Item)

 If PickedItem = "" Then
  Call AddText("There Is No Item In That Inventory Slot!", RED, 9, False, True, False)
   End If
 
 
 If PickedItem = "potion" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Potion", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: A weak potion made of herbs. Heals 25 Health.", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("---", DARKGREY, 9, False, True, False)
         End If

 If PickedItem = "soul potion" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Soul Potion", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: A weak mana potion made with frog feet. Restores 10 Mana", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("---", DARKGREY, 9, False, True, False)
         End If

 If PickedItem = "large potion" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Large Potion", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: A larger jar of a regular potion. Heals 150 Health.", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("---", DARKGREY, 9, False, True, False)
         End If

 If PickedItem = "large soul potion" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Large Soul Potion", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: A larger vial of a regular soul potion. Restores 50 Mana ", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("---", DARKGREY, 9, False, True, False)
         End If
         
 If PickedItem = "copper sword" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Copper Sword", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: A normal sized sword made of copper. Nothing special..", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("Item Stat Bonus: Strength + 3", DARKGREY, 9, False, True, False)
         Call AddText("---", DARKGREY, 9, False, True, False)
          End If
          
 If PickedItem = "hide chestplate" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Hide Chestplate", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: Made of soft hide, not recommended for long term use..", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("Item Stat Bonus: Defence + 2", DARKGREY, 9, False, True, False)
         Call AddText("---", DARKGREY, 9, False, True, False)
          End If

 If PickedItem = "hide leggings" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Hide Leggings", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: Leggings made from soft hide, they smell funny", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("Item Stat Bonus: Defence + 1", DARKGREY, 9, False, True, False)
         Call AddText("---", DARKGREY, 9, False, True, False)
          End If
          
 If PickedItem = "fur gloves" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Fur Gloves", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: Warm and fuzzy gloves made with fake fur, just to make animal lovers happy!", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("Item Stat Bonus: Defence + 1", DARKGREY, 9, False, True, False)
         Call AddText("---", DARKGREY, 9, False, True, False)
          End If
          
 If PickedItem = "fur boots" Then
  Call AddText("---", DARKGREY, 9, False, True, False)
   Call AddText("Fur Boots", BLACK, 9, False, True, False)
    Call AddText("---", DARKGREY, 9, False, True, False)
     Call AddText("", DARKGREY, 9, False, True, False)
      Call AddText("Item Description: This was made with the highest quality leg hairs! Be proud!", DARKGREY, 9, False, True, False)
       Call AddText("", DARKGREY, 9, False, True, False)
        Call AddText("Item Stat Bonus: Defence + 1 / Speed + 1", DARKGREY, 9, False, True, False)
         Call AddText("---", DARKGREY, 9, False, True, False)
          End If
          
          
          
          
          
          
          
          
          

End Sub

Public Sub ItemDescription()



If PText = "/description 1" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem1)
   End If



If PText = "/description 2" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem2)
   End If




If PText = "/description 3" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem3)
   End If




If PText = "/description 4" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem4)
   End If




If PText = "/description 5" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem5)
   End If




If PText = "/description 6" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem6)
   End If




If PText = "/description 7" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem7)
   End If



If PText = "/description 8" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem8)
   End If



If PText = "/description 9" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem9)
   End If



If PText = "/description 10" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem10)
   End If



If PText = "/description 11" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem11)
   End If



If PText = "/description 12" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem12)
   End If



If PText = "/description 13" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem13)
   End If



If PText = "/description 14" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem14)
   End If



If PText = "/description 15" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem15)
   End If



If PText = "/description 16" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem16)
   End If



If PText = "/description 17" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem17)
   End If



If PText = "/description 18" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem18)
   End If



If PText = "/description 19" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem19)
   End If



If PText = "/description 20" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem20)
   End If



If PText = "/description 21" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem21)
   End If



If PText = "/description 22" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem22)
   End If



If PText = "/description 23" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem23)
   End If



If PText = "/description 24" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem24)
   End If



If PText = "/description 25" Then
 Call NullPText
  Call ItemDescriptionDatabase(InventoryItem25)
   End If



End Sub

Public Sub DisplayEquipment()

If PText = "/equips" Then
 Call AddText("", BLACK, 9, False, False, False)
  Call AddText("----------", GREEN, 9, False, True, False)
   Call AddText(CName & "'s Equipment", GREEN, 9, False, True, False)
    Call AddText("----------", GREEN, 9, False, True, False)
     Call AddText("", BLACK, 9, False, False, False)
     Call AddText("Weapon: " & PlayerWeapon, BLACK, 9, False, False, False)
      Call AddText("Helm: " & PlayerHelm, BLACK, 9, False, False, False)
       Call AddText("Armor: " & PlayerArmor, BLACK, 9, False, False, False)
        Call AddText("Gloves: " & PlayerGloves, BLACK, 9, False, False, False)
         Call AddText("Leggings: " & PlayerLeggings, BLACK, 9, False, False, False)
          Call AddText("Boots: " & PlayerBoots, BLACK, 9, False, False, False)
           Call AddText("Accessory: " & PlayerAccessory, BLACK, 9, False, False, False)
            Call AddText("", BLACK, 9, False, False, False)
             End If
           
If PText = "/equipment" Then
 Call AddText("", BLACK, 9, False, False, False)
  Call AddText("----------", GREEN, 9, False, True, False)
   Call AddText(CName & "'s Equipment", GREEN, 9, False, True, False)
    Call AddText("----------", GREEN, 9, False, True, False)
     Call AddText("", BLACK, 9, False, False, False)
     Call AddText("Weapon: " & PlayerWeapon, BLACK, 9, False, False, False)
      Call AddText("Helm: " & PlayerHelm, BLACK, 9, False, False, False)
       Call AddText("Armor: " & PlayerArmor, BLACK, 9, False, False, False)
        Call AddText("Gloves: " & PlayerGloves, BLACK, 9, False, False, False)
         Call AddText("Leggings: " & PlayerLeggings, BLACK, 9, False, False, False)
          Call AddText("Boots: " & PlayerBoots, BLACK, 9, False, False, False)
           Call AddText("Accessory: " & PlayerAccessory, BLACK, 9, False, False, False)
            Call AddText("", BLACK, 9, False, False, False)
             End If


End Sub

Public Sub Equip()

'Equip Item 1
 If PText = "/equip 1" Then
  
  EquippingItem = True
   Call EquipableItemDatabase(InventoryItem1)
   
   
   'check if there is an item in the slot
    If InventoryItem1 = "" Then
     Call AddText("You Do Not Have An Item In That Inventory Slot!", RED, 9, False, False, False)
      Exit Sub
       End If
   '---------
   
   
   
   'check if the item can/cant be equipped
     If WeaponItemEquip = False Then
      If HelmItemEquip = False Then
       If ArmorItemEquip = False Then
        If GlovesItemEquip = False Then
         If LeggingsItemEquip = False Then
          If BootsItemEquip = False Then
           If AccessoryItemEquip = False Then
           
            
            Call AddText("This Item Cannot Be Equipped!", RED, 9, False, False, False)
             Exit Sub
              
             End If
              End If
               End If
                End If
                 End If
                  End If
                   End If
    '-----------------------
    
   'check if item is already equipped // helm to boots
   
   If WeaponItemEquip = True Then
    If PlayerWeapon <> "" Then
     Call AddText("You Already Have An Item Equipped In Your Weapon Slot!", RED, 9, False, False, False)
      Call AddText("Unequip Your Current Weapon To Equip This Item!", RED, 9, False, False, False)
       Exit Sub
        End If
         End If
   
   
    If HelmItemEquip = True Then
     If PlayerHelm <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Helm Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Helm To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
          
    If ArmorItemEquip = True Then
     If PlayerArmor <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Armor Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Armor To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
     
    If GlovesItemEquip = True Then
     If PlayerGloves <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Gloves Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Gloves To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If

    
    If LeggingsItemEquip = True Then
     If PlayerLeggings <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Leggings Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Leggings To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
    If BootsItemEquip = True Then
     If PlayerBoots <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Boots Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Boots To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If


    If AccessoryItemEquip = True Then
     If PlayerAccessory <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Accessory Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Accessory To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
'--------

'finalize the equip, equip it.

If WeaponItemEquip = True Then
 PlayerWeapon = InventoryItem1
End If

If HelmItemEquip = True Then
 PlayerHelm = InventoryItem1
End If

If ArmorItemEquip = True Then
 PlayerArmor = InventoryItem1
End If

If GlovesItemEquip = True Then
 PlayerGloves = InventoryItem1
End If

If LeggingsItemEquip = True Then
 PlayerLeggings = InventoryItem1
End If

If BootsItemEquip = True Then
 PlayerBoots = InventoryItem1
End If

If AccessoryItemEquip = True Then
 PlayerAccessory = InventoryItem1
End If



 Call EquipableItemEffectDatabase(InventoryItem1)
  Call AddText("You Have Equipped The " & InventoryItem1 & "!", GREEN, 9, False, False, False)
   Call SaveGame
 
 End If
 






















'Equip Item 2
 If PText = "/equip 2" Then
  
  EquippingItem = True
   Call EquipableItemDatabase(InventoryItem2)
   
   
   'check if there is an item in the slot
    If InventoryItem2 = "" Then
     Call AddText("You Do Not Have An Item In That Inventory Slot!", RED, 9, False, False, False)
      Exit Sub
       End If
   '---------
   
   
   
   'check if the item can/cant be equipped
     If WeaponItemEquip = False Then
      If HelmItemEquip = False Then
       If ArmorItemEquip = False Then
        If GlovesItemEquip = False Then
         If LeggingsItemEquip = False Then
          If BootsItemEquip = False Then
           If AccessoryItemEquip = False Then
           
            
            Call AddText("This Item Cannot Be Equipped!", RED, 9, False, False, False)
             Exit Sub
              
             End If
              End If
               End If
                End If
                 End If
                  End If
                   End If
    '-----------------------
    
   'check if item is already equipped // helm to boots
   
   If WeaponItemEquip = True Then
    If PlayerWeapon <> "" Then
     Call AddText("You Already Have An Item Equipped In Your Weapon Slot!", RED, 9, False, False, False)
      Call AddText("Unequip Your Current Weapon To Equip This Item!", RED, 9, False, False, False)
       Exit Sub
        End If
         End If
   
   
    If HelmItemEquip = True Then
     If PlayerHelm <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Helm Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Helm To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
          
    If ArmorItemEquip = True Then
     If PlayerArmor <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Armor Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Armor To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
     
    If GlovesItemEquip = True Then
     If PlayerGloves <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Gloves Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Gloves To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If

    
    If LeggingsItemEquip = True Then
     If PlayerLeggings <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Leggings Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Leggings To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
    If BootsItemEquip = True Then
     If PlayerBoots <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Boots Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Boots To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If


    If AccessoryItemEquip = True Then
     If PlayerAccessory <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Accessory Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Accessory To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
'--------

'finalize the equip, equip it.

If WeaponItemEquip = True Then
 PlayerWeapon = InventoryItem2
End If

If HelmItemEquip = True Then
 PlayerHelm = InventoryItem2
End If

If ArmorItemEquip = True Then
 PlayerArmor = InventoryItem2
End If

If GlovesItemEquip = True Then
 PlayerGloves = InventoryItem2
End If

If LeggingsItemEquip = True Then
 PlayerLeggings = InventoryItem2
End If

If BootsItemEquip = True Then
 PlayerBoots = InventoryItem2
End If

If AccessoryItemEquip = True Then
 PlayerAccessory = InventoryItem2
End If



 Call EquipableItemEffectDatabase(InventoryItem2)
  Call AddText("You Have Equipped The " & InventoryItem2 & "!", GREEN, 9, False, False, False)
   Call SaveGame
 
 End If
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 'Equip Item 3
  If PText = "/equip 3" Then
  
  EquippingItem = True
   Call EquipableItemDatabase(InventoryItem3)
   
   
   'check if there is an item in the slot
    If InventoryItem3 = "" Then
     Call AddText("You Do Not Have An Item In That Inventory Slot!", RED, 9, False, False, False)
      Exit Sub
       End If
   '---------
   
   
   
   'check if the item can/cant be equipped
     If WeaponItemEquip = False Then
      If HelmItemEquip = False Then
       If ArmorItemEquip = False Then
        If GlovesItemEquip = False Then
         If LeggingsItemEquip = False Then
          If BootsItemEquip = False Then
           If AccessoryItemEquip = False Then
           
            
            Call AddText("This Item Cannot Be Equipped!", RED, 9, False, False, False)
             Exit Sub
              
             End If
              End If
               End If
                End If
                 End If
                  End If
                   End If
    '-----------------------
    
   'check if item is already equipped // helm to boots
   
   If WeaponItemEquip = True Then
    If PlayerWeapon <> "" Then
     Call AddText("You Already Have An Item Equipped In Your Weapon Slot!", RED, 9, False, False, False)
      Call AddText("Unequip Your Current Weapon To Equip This Item!", RED, 9, False, False, False)
       Exit Sub
        End If
         End If
   
   
    If HelmItemEquip = True Then
     If PlayerHelm <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Helm Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Helm To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
          
    If ArmorItemEquip = True Then
     If PlayerArmor <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Armor Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Armor To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
     
    If GlovesItemEquip = True Then
     If PlayerGloves <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Gloves Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Gloves To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If

    
    If LeggingsItemEquip = True Then
     If PlayerLeggings <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Leggings Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Leggings To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
    If BootsItemEquip = True Then
     If PlayerBoots <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Boots Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Boots To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If


    If AccessoryItemEquip = True Then
     If PlayerAccessory <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Accessory Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Accessory To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
'--------

'finalize the equip, equip it.

If WeaponItemEquip = True Then
 PlayerWeapon = InventoryItem3
End If

If HelmItemEquip = True Then
 PlayerHelm = InventoryItem3
End If

If ArmorItemEquip = True Then
 PlayerArmor = InventoryItem3
End If

If GlovesItemEquip = True Then
 PlayerGloves = InventoryItem3
End If

If LeggingsItemEquip = True Then
 PlayerLeggings = InventoryItem3
End If

If BootsItemEquip = True Then
 PlayerBoots = InventoryItem3
End If

If AccessoryItemEquip = True Then
 PlayerAccessory = InventoryItem3
End If



 Call EquipableItemEffectDatabase(InventoryItem3)
  Call AddText("You Have Equipped The " & InventoryItem3 & "!", GREEN, 9, False, False, False)
   Call SaveGame
 
 End If
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
'Equip Item 4
 If PText = "/equip 4" Then
  
  EquippingItem = True
   Call EquipableItemDatabase(InventoryItem4)
   
   
   'check if there is an item in the slot
    If InventoryItem4 = "" Then
     Call AddText("You Do Not Have An Item In That Inventory Slot!", RED, 9, False, False, False)
      Exit Sub
       End If
   '---------
   
   
   
   'check if the item can/cant be equipped
     If WeaponItemEquip = False Then
      If HelmItemEquip = False Then
       If ArmorItemEquip = False Then
        If GlovesItemEquip = False Then
         If LeggingsItemEquip = False Then
          If BootsItemEquip = False Then
           If AccessoryItemEquip = False Then
           
            
            Call AddText("This Item Cannot Be Equipped!", RED, 9, False, False, False)
             Exit Sub
              
             End If
              End If
               End If
                End If
                 End If
                  End If
                   End If
    '-----------------------
    
   'check if item is already equipped // helm to boots
   
   If WeaponItemEquip = True Then
    If PlayerWeapon <> "" Then
     Call AddText("You Already Have An Item Equipped In Your Weapon Slot!", RED, 9, False, False, False)
      Call AddText("Unequip Your Current Weapon To Equip This Item!", RED, 9, False, False, False)
       Exit Sub
        End If
         End If
   
   
    If HelmItemEquip = True Then
     If PlayerHelm <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Helm Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Helm To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
          
    If ArmorItemEquip = True Then
     If PlayerArmor <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Armor Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Armor To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
     
    If GlovesItemEquip = True Then
     If PlayerGloves <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Gloves Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Gloves To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If

    
    If LeggingsItemEquip = True Then
     If PlayerLeggings <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Leggings Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Leggings To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
    If BootsItemEquip = True Then
     If PlayerBoots <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Boots Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Boots To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If


    If AccessoryItemEquip = True Then
     If PlayerAccessory <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Accessory Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Accessory To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
'--------

'finalize the equip, equip it.

If WeaponItemEquip = True Then
 PlayerWeapon = InventoryItem4
End If

If HelmItemEquip = True Then
 PlayerHelm = InventoryItem4
End If

If ArmorItemEquip = True Then
 PlayerArmor = InventoryItem4
End If

If GlovesItemEquip = True Then
 PlayerGloves = InventoryItem4
End If

If LeggingsItemEquip = True Then
 PlayerLeggings = InventoryItem4
End If

If BootsItemEquip = True Then
 PlayerBoots = InventoryItem4
End If

If AccessoryItemEquip = True Then
 PlayerAccessory = InventoryItem4
End If



 Call EquipableItemEffectDatabase(InventoryItem4)
  Call AddText("You Have Equipped The " & InventoryItem4 & "!", GREEN, 9, False, False, False)
   Call SaveGame
 
 End If
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
'Equip Item 5
 If PText = "/equip 5" Then
  
  EquippingItem = True
   Call EquipableItemDatabase(InventoryItem5)
   
   
   'check if there is an item in the slot
    If InventoryItem5 = "" Then
     Call AddText("You Do Not Have An Item In That Inventory Slot!", RED, 9, False, False, False)
      Exit Sub
       End If
   '---------
   
   
   
   'check if the item can/cant be equipped
     If WeaponItemEquip = False Then
      If HelmItemEquip = False Then
       If ArmorItemEquip = False Then
        If GlovesItemEquip = False Then
         If LeggingsItemEquip = False Then
          If BootsItemEquip = False Then
           If AccessoryItemEquip = False Then
           
            
            Call AddText("This Item Cannot Be Equipped!", RED, 9, False, False, False)
             Exit Sub
              
             End If
              End If
               End If
                End If
                 End If
                  End If
                   End If
    '-----------------------
    
   'check if item is already equipped // helm to boots
   
   If WeaponItemEquip = True Then
    If PlayerWeapon <> "" Then
     Call AddText("You Already Have An Item Equipped In Your Weapon Slot!", RED, 9, False, False, False)
      Call AddText("Unequip Your Current Weapon To Equip This Item!", RED, 9, False, False, False)
       Exit Sub
        End If
         End If
   
   
    If HelmItemEquip = True Then
     If PlayerHelm <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Helm Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Helm To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
          
    If ArmorItemEquip = True Then
     If PlayerArmor <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Armor Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Armor To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
     
    If GlovesItemEquip = True Then
     If PlayerGloves <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Gloves Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Gloves To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If

    
    If LeggingsItemEquip = True Then
     If PlayerLeggings <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Leggings Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Leggings To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
          
    If BootsItemEquip = True Then
     If PlayerBoots <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Boots Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Boots To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If


    If AccessoryItemEquip = True Then
     If PlayerAccessory <> "" Then
      Call AddText("You Already Have An Item Equipped In Your Accessory Slot!", RED, 9, False, False, False)
       Call AddText("Unequip Your Current Accessory To Equip This Item!", RED, 9, False, False, False)
        Exit Sub
         End If
          End If
'--------

'finalize the equip, equip it.

If WeaponItemEquip = True Then
 PlayerWeapon = InventoryItem5
End If

If HelmItemEquip = True Then
 PlayerHelm = InventoryItem5
End If

If ArmorItemEquip = True Then
 PlayerArmor = InventoryItem5
End If

If GlovesItemEquip = True Then
 PlayerGloves = InventoryItem5
End If

If LeggingsItemEquip = True Then
 PlayerLeggings = InventoryItem5
End If

If BootsItemEquip = True Then
 PlayerBoots = InventoryItem5
End If

If AccessoryItemEquip = True Then
 PlayerAccessory = InventoryItem5
End If



 Call EquipableItemEffectDatabase(InventoryItem5)
  Call AddText("You Have Equipped The " & InventoryItem5 & "!", GREEN, 9, False, False, False)
   Call SaveGame
 
 End If
 
 
 
 
 
 
 
 











End Sub

Public Sub EquipableItemDatabase(Item As String)

Dim pItem As String

pItem = LCase(Item)


 
 If pItem = "copper sword" Then
 
 WeaponItemEquip = True
  HelmItemEquip = False
   ArmorItemEquip = False
    GlovesItemEquip = False
     LeggingsItemEquip = False
      BootsItemEquip = False
       AccessoryItemEquip = False

  Exit Sub
 End If




 If pItem = "hide chestplate" Then
 
 WeaponItemEquip = False
  HelmItemEquip = False
   ArmorItemEquip = True
    GlovesItemEquip = False
     LeggingsItemEquip = False
      BootsItemEquip = False
       AccessoryItemEquip = False

  Exit Sub
 End If




 If pItem = "hide leggings" Then
 
 WeaponItemEquip = False
  HelmItemEquip = False
   ArmorItemEquip = False
    GlovesItemEquip = False
     LeggingsItemEquip = True
      BootsItemEquip = False
       AccessoryItemEquip = False

  Exit Sub
 End If




 If pItem = "fur gloves" Then
 
 WeaponItemEquip = False
  HelmItemEquip = False
   ArmorItemEquip = False
    GlovesItemEquip = True
     LeggingsItemEquip = False
      BootsItemEquip = False
       AccessoryItemEquip = False

  Exit Sub
 End If




 If pItem = "fur boots" Then
 
 WeaponItemEquip = False
  HelmItemEquip = False
   ArmorItemEquip = False
    GlovesItemEquip = False
     LeggingsItemEquip = False
      BootsItemEquip = True
       AccessoryItemEquip = False

  Exit Sub
 End If












'---------If not equipable---------------
WeaponItemEquip = False
 HelmItemEquip = False
  ArmorItemEquip = False
   GlovesItemEquip = False
    LeggingsItemEquip = False
     BootsItemEquip = False
      AccessoryItemEquip = False

       
End Sub

Public Sub EquipableItemEffectDatabase(Item As String)

Dim pItem As String

pItem = LCase(Item)


'copper sword
If pItem = "copper sword" Then

 If EquippingItem = True Then
  Strength = Strength + 3
   End If

  If EquippingItem = False Then
   Strength = Strength - 3
    End If

 Exit Sub
End If




'hide chestplate
If pItem = "hide chestplate" Then

 If EquippingItem = True Then
  Defence = Defence + 2
   End If

  If EquippingItem = False Then
   Defence = Defence - 2
    End If

 Exit Sub
End If



'hide leggings
If pItem = "hide chestplate" Then

 If EquippingItem = True Then
  Defence = Defence + 1
   End If

  If EquippingItem = False Then
   Defence = Defence - 1
    End If

 Exit Sub
End If




'fur gloves
If pItem = "fur gloves" Then

 If EquippingItem = True Then
  Defence = Defence + 1
   End If

  If EquippingItem = False Then
   Defence = Defence - 1
    End If

 Exit Sub
End If




'fur boots
If pItem = "fur boots" Then

 If EquippingItem = True Then
  Defence = Defence + 1
   Speed = Speed + 1
   End If

  If EquippingItem = False Then
   Defence = Defence - 1
    Speed = Speed - 1
    End If

 Exit Sub
End If






End Sub

Public Sub UnEquip()

Dim pWeapon As String
Dim pHelm As String
Dim pArmor As String
Dim pGloves As String
Dim pLeggings As String
Dim pBoots As String
Dim pAccessory As String





'Unequip Weapon----------
If PText = "/unequip weapon" Then
 
 
      'if they dont have a helm equipped
        If PlayerWeapon = "" Then
         Call AddText("There Is Not An Item Equipped In Your Weapon Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerWeapon)
            
         pWeapon = PlayerWeapon
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pWeapon)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pWeapon & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If

















'Unequip Helm--------------
If PText = "/unequip helm" Then
 
 
      'if they dont have a helm equipped
        If PlayerHelm = "" Then
         Call AddText("There Is Not An Item Equipped In Your Helm Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerHelm)
            
         pHelm = PlayerHelm
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pHelm)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pHelm & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If
            
            
            
            
            
            
            
            


'Unequip Armor-----------------
If PText = "/unequip armor" Then
 
 
      'if they dont have a helm equipped
        If PlayerArmor = "" Then
         Call AddText("There Is Not An Item Equipped In Your Armor Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerArmor)
            
         pArmor = PlayerArmor
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pArmor)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pArmor & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If
   
   
   
   
   
   
   
   
   
   
   
 'Unequip Gloves------
 If PText = "/unequip gloves" Then
 
 
      'if they dont have a helm equipped
        If PlayerGloves = "" Then
         Call AddText("There Is Not An Item Equipped In Your Gloves Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerGloves)
            
         pGloves = PlayerGloves
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pGloves)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pGloves & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If
   
   
   
   
   
   
   
   
   
   
   
'Unequip Leggings----------
If PText = "/unequip leggings" Then
 
 
      'if they dont have a helm equipped
        If PlayerLeggings = "" Then
         Call AddText("There Is Not An Item Equipped In Your Leggings Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerLeggings)
            
         pLeggings = PlayerLeggings
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pLeggings)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pLeggings & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
'unequip boots--------
If PText = "/unequip boots" Then
 
 
      'if they dont have a helm equipped
        If PlayerBoots = "" Then
         Call AddText("There Is Not An Item Equipped In Your Boots Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerBoots)
            
         pBoots = PlayerBoots
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pBoots)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pBoots & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If
   
   
   
   
   
   
   
   
   
   
   
'unequip accessory
If PText = "/unequip accessory" Then
 
 
      'if they dont have a helm equipped
        If PlayerAccessory = "" Then
         Call AddText("There Is Not An Item Equipped In Your Accessory Slot!", RED, 9, False, False, False)
          Exit Sub
           End If
           
        'check what slots the unequiped item takes
           Call EquipableItemDatabase(PlayerAccessory)
            
         pAccessory = PlayerAccessory
            
            
   'Unequip the slots the item takes
    If WeaponItemEquip = True Then
     PlayerWeapon = ""
      End If
      
    If HelmItemEquip = True Then
     PlayerHelm = ""
      End If
    
    If ArmorItemEquip = True Then
     PlayerArmor = ""
      End If
      
    If GlovesItemEquip = True Then
     PlayerGloves = ""
      End If
      
    If LeggingsItemEquip = True Then
     PlayerLeggings = ""
      End If
      
    If BootsItemEquip = True Then
     PlayerBoots = ""
      End If
    
    If AccessoryItemEquip = True Then
     PlayerAccessory = ""
      End If
      
    '--------
       
       
       'Take stat bonus' off
     EquippingItem = False
      EquipableItemEffectDatabase (pAccessory)
      
            
            
       'display the item has been unequipped
        Call AddText("You Have Unequipped The " & pAccessory & "!", GREEN, 9, False, False, False)
         Call SaveGame
   End If
   



End Sub

Public Function CheckInvWeapon() As Boolean




If InventoryItem1 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem4 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem13 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem14 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem15 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem16 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem17 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem18 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem19 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem20 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem21 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem22 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem23 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem24 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


If InventoryItem25 = PlayerWeapon Then
 CheckInvWeapon = False
  Exit Function
  End If


CheckInvWeapon = True



End Function

Public Function CheckInvHelm() As Boolean



If InventoryItem1 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem4 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem13 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem14 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem15 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem16 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem17 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem18 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem19 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem20 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem21 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem22 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem23 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem24 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


If InventoryItem25 = PlayerHelm Then
 CheckInvHelm = False
  Exit Function
  End If


CheckInvHelm = True



End Function




Public Function CheckInvArmor() As Boolean




If InventoryItem1 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem4 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem13 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem14 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem15 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem16 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem17 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem18 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem19 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem20 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem21 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem22 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem23 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem24 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


If InventoryItem25 = PlayerArmor Then
 CheckInvArmor = False
  Exit Function
  End If


CheckInvArmor = True




End Function


Public Function CheckInvGloves() As Boolean



If InventoryItem1 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem4 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem13 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem14 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem15 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem16 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem17 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem18 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem19 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem20 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem21 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem22 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem23 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem24 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


If InventoryItem25 = PlayerGloves Then
 CheckInvGloves = False
  Exit Function
  End If


CheckInvGloves = True





End Function

Public Function CheckInvLeggings() As Boolean



If InventoryItem1 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem4 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem13 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem14 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem15 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem16 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem17 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem18 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem19 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem20 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem21 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem22 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem23 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem24 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


If InventoryItem25 = PlayerLeggings Then
 CheckInvLeggings = False
  Exit Function
  End If


CheckInvLeggings = True





End Function

Public Function CheckInvBoots() As Boolean



If InventoryItem1 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem4 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem13 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem14 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem15 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem16 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem17 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem18 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem19 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem20 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem21 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem22 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem23 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem24 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


If InventoryItem25 = PlayerBoots Then
 CheckInvBoots = False
  Exit Function
  End If


CheckInvBoots = True




End Function

Public Function CheckInvAccessory() As Boolean





If InventoryItem1 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  

If InventoryItem2 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
  
If InventoryItem3 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem4 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
If InventoryItem5 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  

If InventoryItem6 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
  
If InventoryItem7 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
  
If InventoryItem8 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
  
If InventoryItem9 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
  
If InventoryItem10 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  
  
If InventoryItem11 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If
  

If InventoryItem12 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem13 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem14 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem15 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem16 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem17 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem18 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem19 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem20 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem21 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem22 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem23 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem24 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


If InventoryItem25 = PlayerAccessory Then
 CheckInvAccessory = False
  Exit Function
  End If


CheckInvAccessory = True





End Function

Public Sub CheckSellPrice(Item As String)
  
  Dim pItem As String
   pItem = LCase(Item)

Call CheckBuyPrice(pItem)

ItemSellPrice = ItemBuyPrice / 2


End Sub
