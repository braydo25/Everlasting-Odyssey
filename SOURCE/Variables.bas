Attribute VB_Name = "modVariables"
Option Explicit


'New Character Variables

Public NewSpendable As Integer
Public NewMaxHealth As Integer
Public NewHealth As Integer
Public NewMP As Integer
Public NewMaxMP
Public NewStrength As Integer
Public NewDefence As Integer
Public NewIntelligence As Integer
Public NewSpeed As Integer



'SingleBattle--------------

Public CreatureName As String
Public CreatureHealth As Integer
Public CreatureStrength As Integer
Public CreatureDefence As Integer
Public CreatureExperience As Integer
Public CreatureGold As Integer
Public PlayerMiss As Integer
Public CreatureMiss As Integer
Public CreaturePosion As Boolean
Public CreaturePosionHit As Integer

'Battle Display
Public EnemyDamageTXT
Public PlayerDamageTXT

'MultiBattle----------------


Public MultiCreature1Name As String
Public MultiCreature1Health As Integer
Public MultiCreature1Strength As Integer
Public MultiCreature1Defence As Integer
Public MultiCreature1Experience As Integer
Public MultiCreature1Gold As Integer
Public MultiCreature1Posion As Boolean
Public MultiCreature1PosionHit As Integer
 
Public MultiCreature2Name As String
Public MultiCreature2Health As Integer
Public MultiCreature2Strength As Integer
Public MultiCreature2Defence As Integer
Public MultiCreature2Experience As Integer
Public MultiCreature2Gold As Integer
Public MultiCreature2Posion As Boolean
Public MultiCreature2PosionHit As Integer

'text----------------
Public MyText As String
Public PText As String


'text color----------
Public Const BLACK = 0
Public Const BLUE = 1
Public Const GREEN = 2
Public Const CYAN = 3
Public Const RED = 4
Public Const MAGENTA = 5
Public Const BROWN = 6
Public Const GREY = 7
Public Const DARKGREY = 8
Public Const BRIGHTBLUE = 9
Public Const BRIGHTGREEN = 10
Public Const BRIGHTCYAN = 11
Public Const BRIGHTRED = 12
Public Const PINK = 13
Public Const YELLOW = 14
Public Const WHITE = 15


'Name
Public CName As String


'Stats-----
Public Level As String
Public Experience As String
Public MaxHealth As String
Public Health As String
Public MaxMP As String
Public MP As String
Public Strength As String
Public Defence As String
Public Intelligence As String
Public Speed As String
Public Gold As String
Public StatPoints As String

'Equipped Items
Public PlayerWeapon As String
Public PlayerHelm As String
Public PlayerArmor As String
Public PlayerGloves As String
Public PlayerLeggings As String
Public PlayerBoots As String
Public PlayerAccessory As String

Public EquippingItem As Boolean
Public WeaponItemEquip As Boolean
Public HelmItemEquip As Boolean
Public ArmorItemEquip As Boolean
Public GlovesItemEquip As Boolean
Public LeggingsItemEquip As Boolean
Public BootsItemEquip As Boolean
Public AccessoryItemEquip As Boolean








'Inventory-----------

Public InventoryItem1 As String
Public InventoryItem2 As String
Public InventoryItem3 As String
Public InventoryItem4 As String
Public InventoryItem5 As String
Public InventoryItem6 As String
Public InventoryItem7 As String
Public InventoryItem8 As String
Public InventoryItem9 As String
Public InventoryItem10 As String
Public InventoryItem11 As String
Public InventoryItem12 As String
Public InventoryItem13 As String
Public InventoryItem14 As String
Public InventoryItem15 As String
Public InventoryItem16 As String
Public InventoryItem17 As String
Public InventoryItem18 As String
Public InventoryItem19 As String
Public InventoryItem20 As String
Public InventoryItem21 As String
Public InventoryItem22 As String
Public InventoryItem23 As String
Public InventoryItem24 As String
Public InventoryItem25 As String

'                 ^
'Inventory Amount |
'                 V

Public InventoryItem1Amount As String
Public InventoryItem2Amount As String
Public InventoryItem3Amount As String
Public InventoryItem4Amount As String
Public InventoryItem5Amount As String
Public InventoryItem6Amount As String
Public InventoryItem7Amount As String
Public InventoryItem8Amount As String
Public InventoryItem9Amount As String
Public InventoryItem10Amount As String
Public InventoryItem11Amount As String
Public InventoryItem12Amount As String
Public InventoryItem13Amount As String
Public InventoryItem14Amount As String
Public InventoryItem15Amount As String
Public InventoryItem16Amount As String
Public InventoryItem17Amount As String
Public InventoryItem18Amount As String
Public InventoryItem19Amount As String
Public InventoryItem20Amount As String
Public InventoryItem21Amount As String
Public InventoryItem22Amount As String
Public InventoryItem23Amount As String
Public InventoryItem24Amount As String
Public InventoryItem25Amount As String

'-------------------------------------

'Shop Item Variables
Public ItemBuyPrice As Integer
Public ItemSellPrice As Integer
Public OpenInvSlot As String
Public FullInventory As Boolean

'Item Variables
Public UsableItem As Boolean


'World------
Public Location As String
Public FLocation As String
Public Compass As Boolean
Public ReviveLocation As String


'EVENT FLAGS-----
Public BalsoStartRoom As Integer
Public BalsoRefugeeFight As Integer
Public BalsoRefugeeInFight As Boolean
Public BalsoPortAlura As Integer
Public PeasantJacobQuest As Integer
Public PeasantJacobQuestInFight As Boolean
Public VillageOfGlohovaKingAssistant As Integer

