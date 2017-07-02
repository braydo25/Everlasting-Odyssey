Attribute VB_Name = "modTextInputHandle"
Option Explicit

'Add Text Command
Public Sub AddText(ByVal Msg As String, ByVal color As Integer, size As Long, italic As Boolean, bold As Boolean, under As Boolean)
    frmGAME.txtGame.SelStart = Len(frmGAME.txtGame.Text)
    frmGAME.txtGame.SelFontSize = size
    frmGAME.txtGame.SelItalic = italic
    frmGAME.txtGame.SelBold = bold
    frmGAME.txtGame.SelUnderline = under
    frmGAME.txtGame.SelColor = QBColor(color)
    frmGAME.txtGame.SelText = vbNewLine & (Msg)
    frmGAME.txtGame.SelStart = Len(frmGAME.txtGame.Text) - 1
    
   
End Sub
'---------------



Public Sub TextInput()


'----------------------------------
'------Submit And Write TXT(Player)
'----------------------------------
MyText = frmGAME.txtPlayer.Text
 PText = LCase(frmGAME.txtPlayer.Text)
  frmGAME.txtPlayer.Text = vbNullString
   frmGAME.txtPlayer.SelColor = BLACK
   Call AddText(CName & ">>" & MyText, BLACK, 8, False, False, False)

'----------------------------------
'------Display Full Char Stats-----
'----------------------------------

If PText = "/stats" Then
 Call AddText("Name: " & CName & " || Level: " & Level & " || Experience: " & Experience & " || Health: " & Health & " || Mana: " & MP & " || Strength: " & Strength & " || Defence: " & Defence & " || Intelligence: " & Intelligence & " || Speed: " & Speed & " || Gold: " & Gold & " || Location: " & Location, MAGENTA, 9, False, True, True)
End If



'----------------------------------
'------Check For /Debug------------
'----------------------------------
If PText = "/debug" Then
 frmdebug.Visible = True
End If


'----------------------------------
'------Check For /Help-------------
'----------------------------------

If PText = "/help" Then
 Call AddText("-----", CYAN, 8, False, False, False)
  Call AddText("-Help Menu-", CYAN, 8, False, False, False)
   Call AddText("-----", CYAN, 8, False, False, False)
    Call AddText("", CYAN, 8, False, False, False)
     Call AddText("1. /Commands   -This gives you a list of all the typable commands.", CYAN, 8, False, False, False)
      Call AddText("", CYAN, 8, False, False, False)
     Call AddText("2. You can navigate through towns, outposts, dungeons, and the rest of the world using the North, South, East, West buttons at the appropriate times.", CYAN, 8, False, False, False)
    Call AddText("", CYAN, 8, False, False, False)
   Call AddText("3. To use Area Commands, simply type in the shown command you wish to use and press enter.", CYAN, 8, False, False, False)
  Call AddText("-------------------------", CYAN, 8, False, False, False)
End If
 
'----------------------------------
'------Check For /Save-------------
'----------------------------------

If PText = "/save" Then
 
 If Compass = False Then
  Call AddText("You Cannot Save While Unable To Navigate!", BLUE, 9, False, True, False)
   Exit Sub
    End If


 Call SaveGame
End If

'----------------------------------
'------Check For /Map--------------
'----------------------------------

If PText = "/map" Then
 If frmWorldMap.Visible = False Then
  frmWorldMap.Visible = True
   Call NullPText
 End If
End If

If PText = "/map" Then
 If frmWorldMap.Visible = True Then
  frmWorldMap.Visible = False
   Call NullPText
 End If
End If

'----------------------------------
'------Check For /Inventory--------
'----------------------------------

If PText = "/inventory" Then
 Call NullPText
  Call DisplayInventory
   End If

'----------------------------------
'------Check For Dropping An Item--
'----------------------------------

Call DropItem

'----------------------------------
'------Check For Using An Item-----
'----------------------------------

Call UseItem

'----------------------------------
'------Check For /Commands---------
'----------------------------------

If PText = "/commands" Then
 Call AddText("", CYAN, 9, False, False, False)
  Call AddText("----------", CYAN, 9, False, False, False)
   Call AddText("Commands", CYAN, 9, False, False, False)
    Call AddText("----------", CYAN, 9, False, False, False)
     Call AddText("", CYAN, 9, False, False, False)
      Call AddText("1. /Help : This Brings Up The Help Menu.", CYAN, 9, False, False, False)
       Call AddText("2. /Stats : This Shows All Of Your Stats Including Experience, Gold, Etc", CYAN, 9, False, False, False)
        Call AddText("3. /Save : This Saves Your Game Manually (Autosave Is Enabled)", CYAN, 9, False, False, False)
         Call AddText("4. /Inventory : This Displays Your Inventory", CYAN, 9, False, False, False)
          Call AddText("5. /Use (Item Number) : This Uses An Item From The Chosen Item Slot", CYAN, 9, False, False, False)
           Call AddText("6. /Description (Item Number) : This Gives The Description Of An Item", CYAN, 9, False, False, False)
            Call AddText("7. /Map : This Displays Or Stops Displaying The World Map", CYAN, 9, False, False, False)
             Call AddText("8. /Debug : This Displays A Small Debug Menu For Developers", CYAN, 9, False, False, False)
             End If
      
'----------------------------------
'------Check For Item Descripion--
'----------------------------------

Call ItemDescription

'----------------------------------
'------Check For Equipment---------
'----------------------------------

Call DisplayEquipment

'----------------------------------
'------Check For /Equip------------
'----------------------------------

Call Equip

'----------------------------------
'------Check For /Unequip----------
'----------------------------------

Call UnEquip





End Sub


Public Sub NullPText()

PText = ""

End Sub

