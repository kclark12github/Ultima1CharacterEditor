Public Class Player
    Private mFolderName As String = "C:\Program Files (x86)\Electronic Arts\Ultima 1\Game\Game" 'TODO: Make a registry entry...
    Public Sub New(ByVal PlayerNum As Int16)
        mPlayerNumber = PlayerNum
        Read()
    End Sub
#Region "Properties"
#Region "Enumerations"
    Public Enum enumArmor As Int16
        Skin = 0
        Leather = 1
        ChainMail = 2
        PlateMail = 3
        VacuumSuit = 4
        ReflectSuit = 5
    End Enum
    Public Enum enumClass As Int16
        Fighter = 0
        Cleric = 1
        Wizard = 2
        Thief = 3
    End Enum
    Public Enum enumRace As Int16
        Human = 0
        Elf = 1
        Dwarf = 2
        Bobbit = 3
    End Enum
    Public Enum enumSex As Int16
        Male = 0
        Female = 1
    End Enum
    Public Enum enumSpell As Int16
        Open = 0
        Unlock = 1
        MagicMissile = 2
        Steal = 3
        LadderDown = 4
        LadderUp = 5
        Blink = 6
        Create = 7
        Destroy = 8
        Kill = 9
    End Enum
    Public Enum enumTransport As Int16
        Foot = 0
        Horse = 1
        Cart = 2
        Raft = 3
        Frigate = 4
        Aircar = 5
        Shuttle = 6
        TimeMachine = 7
    End Enum
    Public Enum enumWeapon As Int16
        Hands = 0
        Dagger = 1
        Mace = 2
        Axe = 3
        RopeAndSpikes = 4
        Sword = 5
        GreatSword = 6
        BowAndArrows = 7
        Amulet = 8
        Wand = 9
        Staff = 10
        Triangle = 11
        Pistol = 12
        LightSword = 13
        Phazor = 14
        Blaster = 15
    End Enum
#End Region
#Region "Declarations"
    Public maxAttribute As Int16 = 30
    Public maxItems As Int16 = 9999
    Public maxValue As Int16 = 9999

    Private mChanged As Boolean = False
    Private mPlayerNumber As Short = 0
    Private mReading As Boolean = False

    Private mFileName As String = ""
    Private mName As String = ""
    Private mRace As enumRace = enumRace.Human
    Private mClass As enumClass = enumClass.Fighter
    Private mSex As enumSex = enumSex.Male
    Private mHits As Int16 = 0
    Private mStrength As Int16 = 0
    Private mAgility As Int16 = 0
    Private mStamina As Int16 = 0
    Private mCharisma As Int16 = 0
    Private mWisdom As Int16 = 0
    Private mIntelligence As Int16 = 0
    Private mCoin As Int16 = 0
    Private mExperience As Int16 = 0
    Private mFood As Int16 = 0
    Private mReadyWeapon As enumWeapon = enumWeapon.Hands
    Private mReadySpell As enumSpell = 0
    Private mReadyArmor As enumArmor = enumArmor.Skin
    Private mTransport As enumTransport = enumTransport.Foot
    Private m32h As Int16 = 0
    Private mXlocation As Int16 = 0
    Private mYlocation As Int16 = 0
    Private m38h As Int16 = 0
    Private m3Ah As Int16 = 0
    Private m3Ch As Int16 = 0
    Private m3Eh As Int16 = 0
    Private m40h As Int16 = 0
    Private m42h As Int16 = 0
    Private m44h As Int16 = 0   'FF FF
    Private m46h As Int16 = 0
    Private m48h As Int16 = 0
    Private m4Ah As Int16 = 0   'FF FF
    Private mRedGem As Int16 = 0
    Private mGreenGem As Int16 = 0
    Private mBlueGem As Int16 = 0
    Private mWhiteGem As Int16 = 0
    Private m54h As Int16 = 0   'FF FF
    Private mArmorLeather As Int16 = 0
    Private mArmorChainMail As Int16 = 0
    Private mArmorPlateMail As Int16 = 0
    Private mArmorVacuumSuit As Int16 = 0
    Private mArmorReflectSuit As Int16 = 0
    Private m60h As Int16 = 0   'FF FF
    Private mWeaponDagger As Int16 = 0
    Private mWeaponMace As Int16 = 0
    Private mWeaponAxe As Int16 = 0
    Private mWeaponRopeAndSpikes As Int16 = 0
    Private mWeaponSword As Int16 = 0
    Private mWeaponGreatSword As Int16 = 0
    Private mWeaponBowAndArrows As Int16 = 0
    Private mWeaponAmulet As Int16 = 0
    Private mWeaponWand As Int16 = 0
    Private mWeaponStaff As Int16 = 0
    Private mWeaponTriangle As Int16 = 0
    Private mWeaponPistol As Int16 = 0
    Private mWeaponLightSword As Int16 = 0
    Private mWeaponPhazor As Int16 = 0
    Private mWeaponBlaster As Int16 = 0
    Private m80h As Int16 = 0   'FF FF
    Private mSpellOpen As Int16 = 0
    Private mSpellUnlock As Int16 = 0
    Private mSpellMagicMissile As Int16 = 0
    Private mSpellSteal As Int16 = 0
    Private mSpellLadderDown As Int16 = 0
    Private mSpellLadderUp As Int16 = 0
    Private mSpellBlink As Int16 = 0
    Private mSpellCreate As Int16 = 0
    Private mSpellDestroy As Int16 = 0
    Private mSpellKill As Int16 = 0
    Private m96h As Int16 = 0   'FF FF
    Private mTransportHorse As Int16 = 0
    Private mTransportCart As Int16 = 0
    Private mTransportRaft As Int16 = 0
    Private mTransportFrigate As Int16 = 0
    Private mTransportAircar As Int16 = 0
    Private mTransportShuttle As Int16 = 0
    Private mTransportTimeMachine As Int16 = 0
    Private mA6h As Int16 = 0   'FF FF
    Private mA8h As Int16 = 0   'FF FF
    Private mAAh As Int16 = 0   'FF FF
    Private mMoveCount As Int16 = 0
#End Region
    Public Property [Class] As enumClass
        Get
            Return mClass
        End Get
        Set(value As enumClass)
            Select Case value
                Case enumClass.Fighter, enumClass.Cleric, enumClass.Wizard, enumClass.Thief
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mClass <> value Then mClass = value : SetChanged(True)
        End Set
    End Property
    Public ReadOnly Property Changed As Boolean
        Get
            Return mChanged
        End Get
    End Property
    Public ReadOnly Property FileName As String
        Get
            Return mFileName
        End Get
    End Property
    Public Property Coin As Int16
        Get
            Return mCoin
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mCoin <> value Then mCoin = value : SetChanged(True)
        End Set
    End Property
    Public Property Experience As Int16
        Get
            Return mExperience
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mExperience <> value Then mExperience = value : SetChanged(True)
        End Set
    End Property
    Public Property Food As Int16
        Get
            Return mFood
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mFood <> value Then mFood = value : SetChanged(True)
        End Set
    End Property
    Public Property Hits As Int16
        Get
            Return mHits
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mHits <> value Then mHits = value : SetChanged(True)
        End Set
    End Property
    Public ReadOnly Property Location As String
        Get
            Return String.Format("X={0}, Y={1}", mXlocation, mYlocation)
        End Get
    End Property
    Public ReadOnly Property MoveCount As Int16
        Get
            Return mMoveCount
        End Get
    End Property
    Public ReadOnly Property Name As String
        Get
            Return mName
        End Get
    End Property
    Public Property Race As enumRace
        Get
            Return mRace
        End Get
        Set(value As enumRace)
            Select Case value
                Case enumRace.Human, enumRace.Elf, enumRace.Dwarf, enumRace.Bobbit
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mRace <> value Then mRace = value : SetChanged(True)
        End Set
    End Property
    Public Property Sex As enumSex
        Get
            Return mSex
        End Get
        Set(value As enumSex)
            Select Case value
                Case enumSex.Male, enumSex.Female
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mSex <> value Then mSex = value : SetChanged(True)
        End Set
    End Property
    Public Property Transport As enumTransport
        Get
            Return mTransport
        End Get
        Set(value As enumTransport)
            Select Case value
                Case enumTransport.Aircar, enumTransport.Cart, enumTransport.Foot, enumTransport.Frigate,
                     enumTransport.Horse, enumTransport.Raft, enumTransport.Shuttle, enumTransport.TimeMachine
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mTransport <> value Then mTransport = value : SetChanged(True)
        End Set
    End Property
#Region "Armor"
    Public Property ArmorChainMail As Int16
        Get
            Return mArmorChainMail
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mArmorChainMail <> value Then mArmorChainMail = value : SetChanged(True)
        End Set
    End Property
    Public Property ArmorLeather As Int16
        Get
            Return mArmorLeather
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mArmorLeather <> value Then mArmorLeather = value : SetChanged(True)
        End Set
    End Property
    Public Property ArmorPlateMail As Int16
        Get
            Return mArmorPlateMail
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mArmorPlateMail <> value Then mArmorPlateMail = value : SetChanged(True)
        End Set
    End Property
    Public Property ArmorReflectSuit As Int16
        Get
            Return mArmorReflectSuit
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mArmorReflectSuit <> value Then mArmorReflectSuit = value : SetChanged(True)
        End Set
    End Property
    Public Property ArmorVacuumSuit As Int16
        Get
            Return mArmorVacuumSuit
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mArmorVacuumSuit <> value Then mArmorVacuumSuit = value : SetChanged(True)
        End Set
    End Property
#End Region
#Region "Attributes"
    Public Property Strength As Int16
        Get
            Return mStrength
        End Get
        Set(value As Int16)
            If value > maxAttribute Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxAttribute))
            If mStrength <> value Then mStrength = value : SetChanged(True)
        End Set
    End Property
    Public Property Agility As Int16
        Get
            Return mAgility
        End Get
        Set(value As Int16)
            If value > maxAttribute Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxAttribute))
            If mAgility <> value Then mAgility = value : SetChanged(True)
        End Set
    End Property
    Public Property Stamina As Int16
        Get
            Return mStamina
        End Get
        Set(value As Int16)
            If value > maxAttribute Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxAttribute))
            If mStamina <> value Then mStamina = value : SetChanged(True)
        End Set
    End Property
    Public Property Charisma As Int16
        Get
            Return mCharisma
        End Get
        Set(value As Int16)
            If value > maxAttribute Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxAttribute))
            If mCharisma <> value Then mCharisma = value : SetChanged(True)
        End Set
    End Property
    Public Property Wisdom As Int16
        Get
            Return mWisdom
        End Get
        Set(value As Int16)
            If value > maxAttribute Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxAttribute))
            If mWisdom <> value Then mWisdom = value : SetChanged(True)
        End Set
    End Property
    Public Property Intelligence As Int16
        Get
            Return mIntelligence
        End Get
        Set(value As Int16)
            If value > maxAttribute Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxAttribute))
            If mIntelligence <> value Then mIntelligence = value : SetChanged(True)
        End Set
    End Property
#End Region
#Region "Gems"
    Public Property BlueGem As Int16
        Get
            Return mBlueGem
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mBlueGem <> value Then mBlueGem = value : SetChanged(True)
        End Set
    End Property
    Public Property GreenGem As Int16
        Get
            Return mGreenGem
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mGreenGem <> value Then mGreenGem = value : SetChanged(True)
        End Set
    End Property
    Public Property RedGem As Int16
        Get
            Return mRedGem
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mRedGem <> value Then mRedGem = value : SetChanged(True)
        End Set
    End Property
    Public Property WhiteGem As Int16
        Get
            Return mWhiteGem
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWhiteGem <> value Then mWhiteGem = value : SetChanged(True)
        End Set
    End Property
#End Region
#Region "Ready Items"
    Public Property ReadyArmor As enumArmor
        Get
            Return mReadyArmor
        End Get
        Set(value As enumArmor)
            Select Case value
                Case enumArmor.ChainMail, enumArmor.Leather, enumArmor.PlateMail, enumArmor.ReflectSuit,
                     enumArmor.Skin, enumArmor.VacuumSuit
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mReadyArmor <> value Then mReadyArmor = value : SetChanged(True)
        End Set
    End Property
    Public Property ReadySpell As enumSpell
        Get
            Return mReadySpell
        End Get
        Set(value As enumSpell)
            Select Case value
                Case enumSpell.Blink, enumSpell.Create, enumSpell.Destroy, enumSpell.Kill, enumSpell.LadderDown,
                     enumSpell.LadderUp, enumSpell.MagicMissile, enumSpell.Open, enumSpell.Steal, enumSpell.Unlock
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mReadySpell <> value Then mReadySpell = value : SetChanged(True)
        End Set
    End Property
    Public Property ReadyWeapon As enumWeapon
        Get
            Return mReadyWeapon
        End Get
        Set(value As enumWeapon)
            Select Case value
                Case enumWeapon.Amulet, enumWeapon.Axe, enumWeapon.Blaster, enumWeapon.BowAndArrows, enumWeapon.Dagger, enumWeapon.GreatSword,
                     enumWeapon.Hands, enumWeapon.LightSword, enumWeapon.Mace, enumWeapon.Phazor, enumWeapon.Pistol, enumWeapon.RopeAndSpikes,
                     enumWeapon.Staff, enumWeapon.Sword, enumWeapon.Triangle, enumWeapon.Wand
                Case Else : Throw New ArgumentException(String.Format("Invalid property value ({0})", value))
            End Select
            If mReadyWeapon <> value Then mReadyWeapon = value : SetChanged(True)
        End Set
    End Property
#End Region
#Region "Spells"
    Public Property SpellBlink As Int16
        Get
            Return mSpellBlink
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellBlink <> value Then mSpellBlink = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellCreate As Int16
        Get
            Return mSpellCreate
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellCreate <> value Then mSpellCreate = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellDestroy As Int16
        Get
            Return mSpellDestroy
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellDestroy <> value Then mSpellDestroy = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellKill As Int16
        Get
            Return mSpellKill
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellKill <> value Then mSpellKill = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellLadderDown As Int16
        Get
            Return mSpellLadderDown
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellLadderDown <> value Then mSpellLadderDown = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellLadderUp As Int16
        Get
            Return mSpellLadderUp
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellLadderUp <> value Then mSpellLadderUp = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellMagicMissile As Int16
        Get
            Return mSpellMagicMissile
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellMagicMissile <> value Then mSpellMagicMissile = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellOpen As Int16
        Get
            Return mSpellOpen
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellOpen <> value Then mSpellOpen = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellSteal As Int16
        Get
            Return mSpellSteal
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellSteal <> value Then mSpellSteal = value : SetChanged(True)
        End Set
    End Property
    Public Property SpellUnlock As Int16
        Get
            Return mSpellUnlock
        End Get
        Set(value As Int16)
            If value > maxValue Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxValue))
            If mSpellUnlock <> value Then mSpellUnlock = value : SetChanged(True)
        End Set
    End Property
#End Region
#Region "Transports"
    Public Property TransportAircar As Int16
        Get
            Return mTransportAircar
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportAircar <> value Then mTransportAircar = value : SetChanged(True)
        End Set
    End Property
    Public Property TransportCart As Int16
        Get
            Return mTransportCart
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportCart <> value Then mTransportCart = value : SetChanged(True)
        End Set
    End Property
    Public Property TransportFrigate As Int16
        Get
            Return mTransportFrigate
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportFrigate <> value Then mTransportFrigate = value : SetChanged(True)
        End Set
    End Property
    Public Property TransportHorse As Int16
        Get
            Return mTransportHorse
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportHorse <> value Then mTransportHorse = value : SetChanged(True)
        End Set
    End Property
    Public Property TransportRaft As Int16
        Get
            Return mTransportRaft
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportRaft <> value Then mTransportRaft = value : SetChanged(True)
        End Set
    End Property
    Public Property TransportShuttle As Int16
        Get
            Return mTransportShuttle
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportShuttle <> value Then mTransportShuttle = value : SetChanged(True)
        End Set
    End Property
    Public Property TransportTimeMachine As Int16
        Get
            Return mTransportTimeMachine
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mTransportTimeMachine <> value Then mTransportTimeMachine = value : SetChanged(True)
        End Set
    End Property
#End Region
#Region "Weapons"
    Public Property WeaponAmulet As Int16
        Get
            Return mWeaponAmulet
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponAmulet <> value Then mWeaponAmulet = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponAxe As Int16
        Get
            Return mWeaponAxe
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponAxe <> value Then mWeaponAxe = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponBlaster As Int16
        Get
            Return mWeaponBlaster
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponBlaster <> value Then mWeaponBlaster = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponBowAndArrows As Int16
        Get
            Return mWeaponBowAndArrows
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponBowAndArrows <> value Then mWeaponBowAndArrows = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponDagger As Int16
        Get
            Return mWeaponDagger
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponDagger <> value Then mWeaponDagger = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponGreatSword As Int16
        Get
            Return mWeaponGreatSword
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponGreatSword <> value Then mWeaponGreatSword = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponLightSword As Int16
        Get
            Return mWeaponLightSword
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponLightSword <> value Then mWeaponLightSword = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponMace As Int16
        Get
            Return mWeaponMace
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponMace <> value Then mWeaponMace = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponPhazor As Int16
        Get
            Return mWeaponPhazor
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponPhazor <> value Then mWeaponPhazor = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponPistol As Int16
        Get
            Return mWeaponPistol
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponPistol <> value Then mWeaponPistol = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponRopeAndSpikes As Int16
        Get
            Return mWeaponRopeAndSpikes
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponRopeAndSpikes <> value Then mWeaponRopeAndSpikes = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponStaff As Int16
        Get
            Return mWeaponStaff
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponStaff <> value Then mWeaponStaff = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponSword As Int16
        Get
            Return mWeaponSword
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponSword <> value Then mWeaponSword = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponTriangle As Int16
        Get
            Return mWeaponTriangle
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponTriangle <> value Then mWeaponTriangle = value : SetChanged(True)
        End Set
    End Property
    Public Property WeaponWand As Int16
        Get
            Return mWeaponWand
        End Get
        Set(value As Int16)
            If value > maxItems Then Throw New ArgumentException(String.Format("Property ({0}) cannot be greater than {1}", value, maxItems))
            If mWeaponWand <> value Then mWeaponWand = value : SetChanged(True)
        End Set
    End Property
#End Region
#End Region
#Region "Methods"
    Private Sub Backup()
        Dim fi As FileInfo = New FileInfo(mFileName)
        Dim backup As String = fi.Name.Replace(fi.Extension, String.Format(".{0:yyyyMMdd.HHmmssff}{1}", fi.LastWriteTime, fi.Extension))
        Dim backupPath As String = String.Format("{0}\{1}", fi.DirectoryName, backup)
        If Not FileIO.FileSystem.FileExists(backupPath) Then
            FileIO.FileSystem.RenameFile(mFileName, backup)
            FileIO.FileSystem.CopyFile(backupPath, mFileName, FileIO.UIOption.OnlyErrorDialogs)
        End If
    End Sub
    Private Function ConvertFromChar16(ByVal asciiChars As Char()) As String
        ConvertFromChar16 = ""
        For i As Short = 0 To 15
            If asciiChars(i) = vbNullChar Then Exit For
            ConvertFromChar16 &= asciiChars(i)
        Next i
    End Function
    Private Function ConvertToBytes16(ByVal unicodeString As String) As Byte()
        Dim ascii As Encoding = Encoding.ASCII
        Dim unicode As Encoding = Encoding.Unicode

        Dim unicodeBytes As Byte() = unicode.GetBytes(unicodeString) 'Convert the string into a byte array.
        'Perform the conversion from one encoding to the other.
        Dim asciiBytes As Byte() = Encoding.Convert(unicode, ascii, unicodeBytes)
        ReDim Preserve asciiBytes(15)
        For i As Short = unicodeString.Length To 15 : asciiBytes(i) = 0 : Next i
        Return asciiBytes
    End Function
    Private Sub SetChanged(ByVal value As Boolean)
        If Not mReading Then mChanged = value
    End Sub
    Public Sub Read()
        Dim binReader As BinaryReader = Nothing
        Try
            mFileName = String.Format("{0}\PLAYER{1}.U1", mFolderName, mPlayerNumber)
            If Not File.Exists(mFileName) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", mFileName))
            Dim FileLengthInBytes As Integer = FileLen(mFileName)
            mReading = True

            Dim offset As Integer = 0
            binReader = New BinaryReader(File.Open(mFileName, FileMode.Open))
            mName = ConvertFromChar16(binReader.ReadChars(16)) : offset += 16
            Me.Race = binReader.ReadInt16() : offset += 2
            Me.Class = binReader.ReadInt16() : offset += 2
            Me.Sex = binReader.ReadInt16() : offset += 2
            Me.Hits = binReader.ReadInt16() : offset += 2
            mStrength = binReader.ReadInt16() : offset += 2
            mAgility = binReader.ReadInt16() : offset += 2
            mStamina = binReader.ReadInt16() : offset += 2
            mCharisma = binReader.ReadInt16() : offset += 2
            mWisdom = binReader.ReadInt16() : offset += 2
            mIntelligence = binReader.ReadInt16() : offset += 2
            Me.Coin = binReader.ReadInt16() : offset += 2
            Me.Experience = binReader.ReadInt16() : offset += 2
            Me.Food = binReader.ReadInt16() : offset += 2
            Me.ReadyWeapon = binReader.ReadInt16() : offset += 2
            Me.ReadySpell = binReader.ReadInt16() : offset += 2
            Me.ReadyArmor = binReader.ReadInt16() : offset += 2
            Me.Transport = binReader.ReadInt16() : offset += 2
            m32h = binReader.ReadInt16() : offset += 2
            mXlocation = binReader.ReadInt16() : offset += 2
            mYlocation = binReader.ReadInt16() : offset += 2
            m38h = binReader.ReadInt16() : offset += 2
            m3Ah = binReader.ReadInt16() : offset += 2
            m3Ch = binReader.ReadInt16() : offset += 2
            m3Eh = binReader.ReadInt16() : offset += 2
            m40h = binReader.ReadInt16() : offset += 2
            m42h = binReader.ReadInt16() : offset += 2
            m44h = binReader.ReadInt16() : offset += 2   'FF FF
            m46h = binReader.ReadInt16() : offset += 2
            m48h = binReader.ReadInt16() : offset += 2
            m4Ah = binReader.ReadInt16() : offset += 2   'FF FF
            Me.RedGem = binReader.ReadInt16() : offset += 2
            Me.GreenGem = binReader.ReadInt16() : offset += 2
            Me.BlueGem = binReader.ReadInt16() : offset += 2
            Me.WhiteGem = binReader.ReadInt16() : offset += 2
            m54h = binReader.ReadInt16() : offset += 2   'FF FF
            Me.ArmorLeather = binReader.ReadInt16() : offset += 2
            Me.ArmorChainMail = binReader.ReadInt16() : offset += 2
            Me.ArmorPlateMail = binReader.ReadInt16() : offset += 2
            Me.ArmorVacuumSuit = binReader.ReadInt16() : offset += 2
            Me.ArmorReflectSuit = binReader.ReadInt16() : offset += 2
            m60h = binReader.ReadInt16() : offset += 2   'FF FF
            Me.WeaponDagger = binReader.ReadInt16() : offset += 2
            Me.WeaponMace = binReader.ReadInt16() : offset += 2
            Me.WeaponAxe = binReader.ReadInt16() : offset += 2
            Me.WeaponRopeAndSpikes = binReader.ReadInt16() : offset += 2
            Me.WeaponSword = binReader.ReadInt16() : offset += 2
            Me.WeaponGreatSword = binReader.ReadInt16() : offset += 2
            Me.WeaponBowAndArrows = binReader.ReadInt16() : offset += 2
            Me.WeaponAmulet = binReader.ReadInt16() : offset += 2
            Me.WeaponWand = binReader.ReadInt16() : offset += 2
            Me.WeaponStaff = binReader.ReadInt16() : offset += 2
            Me.WeaponTriangle = binReader.ReadInt16() : offset += 2
            Me.WeaponPistol = binReader.ReadInt16() : offset += 2
            Me.WeaponLightSword = binReader.ReadInt16() : offset += 2
            Me.WeaponPhazor = binReader.ReadInt16() : offset += 2
            Me.WeaponBlaster = binReader.ReadInt16() : offset += 2
            m80h = binReader.ReadInt16() : offset += 2   'FF FF
            Me.SpellOpen = binReader.ReadInt16() : offset += 2
            Me.SpellUnlock = binReader.ReadInt16() : offset += 2
            Me.SpellMagicMissile = binReader.ReadInt16() : offset += 2
            Me.SpellSteal = binReader.ReadInt16() : offset += 2
            Me.SpellLadderDown = binReader.ReadInt16() : offset += 2
            Me.SpellLadderUp = binReader.ReadInt16() : offset += 2
            Me.SpellBlink = binReader.ReadInt16() : offset += 2
            Me.SpellCreate = binReader.ReadInt16() : offset += 2
            Me.SpellDestroy = binReader.ReadInt16() : offset += 2
            Me.SpellKill = binReader.ReadInt16() : offset += 2
            m96h = binReader.ReadInt16() : offset += 2   'FF FF
            Me.TransportHorse = binReader.ReadInt16() : offset += 2
            Me.TransportCart = binReader.ReadInt16() : offset += 2
            Me.TransportRaft = binReader.ReadInt16() : offset += 2
            Me.TransportFrigate = binReader.ReadInt16() : offset += 2
            Me.TransportAircar = binReader.ReadInt16() : offset += 2
            Me.TransportShuttle = binReader.ReadInt16() : offset += 2
            Me.TransportTimeMachine = binReader.ReadInt16() : offset += 2
            mA6h = binReader.ReadInt16() : offset += 2   'Enemy Vessels
            mA8h = binReader.ReadInt16() : offset += 2   'Sign Marker
            mAAh = binReader.ReadInt16() : offset += 2
            mMoveCount = binReader.ReadInt16() : offset += 2
            mReading = False
            SetChanged(False)
        Finally
            'FileSystem.FileClose(FileUnit)
            If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
        End Try
    End Sub
    Public Sub Save()
        Dim binWriter As BinaryWriter = Nothing
        Try
            Backup()

            Dim offset As Integer = 0
            Dim buffer() As Byte = ConvertToBytes16(mName)

            binWriter = New BinaryWriter(File.Open(mFileName, FileMode.Open))
            binWriter.Write(buffer) : offset += 16
            binWriter.Write(Me.Race) : offset += 2
            binWriter.Write(Me.Class) : offset += 2
            binWriter.Write(Me.Sex) : offset += 2
            binWriter.Write(Me.Hits) : offset += 2
            binWriter.Write(mStrength) : offset += 2
            binWriter.Write(mAgility) : offset += 2
            binWriter.Write(mStamina) : offset += 2
            binWriter.Write(mCharisma) : offset += 2
            binWriter.Write(mWisdom) : offset += 2
            binWriter.Write(mIntelligence) : offset += 2
            binWriter.Write(Me.Coin) : offset += 2
            binWriter.Write(Me.Experience) : offset += 2
            binWriter.Write(Me.Food) : offset += 2
            binWriter.Write(Me.ReadyWeapon) : offset += 2
            binWriter.Write(Me.ReadySpell) : offset += 2
            binWriter.Write(Me.ReadyArmor) : offset += 2
            binWriter.Write(Me.Transport) : offset += 2
            binWriter.Write(m32h) : offset += 2
            binWriter.Write(mXlocation) : offset += 2
            binWriter.Write(mYlocation) : offset += 2
            binWriter.Write(m38h) : offset += 2
            binWriter.Write(m3Ah) : offset += 2
            binWriter.Write(m3Ch) : offset += 2
            binWriter.Write(m3Eh) : offset += 2
            binWriter.Write(m40h) : offset += 2
            binWriter.Write(m42h) : offset += 2
            binWriter.Write(m44h) : offset += 2   'FF FF
            binWriter.Write(m46h) : offset += 2
            binWriter.Write(m48h) : offset += 2
            binWriter.Write(m4Ah) : offset += 2   'FF FF
            binWriter.Write(Me.RedGem) : offset += 2
            binWriter.Write(Me.GreenGem) : offset += 2
            binWriter.Write(Me.BlueGem) : offset += 2
            binWriter.Write(Me.WhiteGem) : offset += 2
            binWriter.Write(m54h) : offset += 2   'FF FF
            binWriter.Write(Me.ArmorLeather) : offset += 2
            binWriter.Write(Me.ArmorChainMail) : offset += 2
            binWriter.Write(Me.ArmorPlateMail) : offset += 2
            binWriter.Write(Me.ArmorVacuumSuit) : offset += 2
            binWriter.Write(Me.ArmorReflectSuit) : offset += 2
            binWriter.Write(m60h) : offset += 2   'FF FF
            binWriter.Write(Me.WeaponDagger) : offset += 2
            binWriter.Write(Me.WeaponMace) : offset += 2
            binWriter.Write(Me.WeaponAxe) : offset += 2
            binWriter.Write(Me.WeaponRopeAndSpikes) : offset += 2
            binWriter.Write(Me.WeaponSword) : offset += 2
            binWriter.Write(Me.WeaponGreatSword) : offset += 2
            binWriter.Write(Me.WeaponBowAndArrows) : offset += 2
            binWriter.Write(Me.WeaponAmulet) : offset += 2
            binWriter.Write(Me.WeaponWand) : offset += 2
            binWriter.Write(Me.WeaponStaff) : offset += 2
            binWriter.Write(Me.WeaponTriangle) : offset += 2
            binWriter.Write(Me.WeaponPistol) : offset += 2
            binWriter.Write(Me.WeaponLightSword) : offset += 2
            binWriter.Write(Me.WeaponPhazor) : offset += 2
            binWriter.Write(Me.WeaponBlaster) : offset += 2
            binWriter.Write(m80h) : offset += 2   'FF FF
            binWriter.Write(Me.SpellOpen) : offset += 2
            binWriter.Write(Me.SpellUnlock) : offset += 2
            binWriter.Write(Me.SpellMagicMissile) : offset += 2
            binWriter.Write(Me.SpellSteal) : offset += 2
            binWriter.Write(Me.SpellLadderDown) : offset += 2
            binWriter.Write(Me.SpellLadderUp) : offset += 2
            binWriter.Write(Me.SpellBlink) : offset += 2
            binWriter.Write(Me.SpellCreate) : offset += 2
            binWriter.Write(Me.SpellDestroy) : offset += 2
            binWriter.Write(Me.SpellKill) : offset += 2
            binWriter.Write(m96h) : offset += 2   'FF FF
            binWriter.Write(Me.TransportHorse) : offset += 2
            binWriter.Write(Me.TransportCart) : offset += 2
            binWriter.Write(Me.TransportRaft) : offset += 2
            binWriter.Write(Me.TransportFrigate) : offset += 2
            binWriter.Write(Me.TransportAircar) : offset += 2
            binWriter.Write(Me.TransportShuttle) : offset += 2
            binWriter.Write(Me.TransportTimeMachine) : offset += 2
            binWriter.Write(mA6h) : offset += 2   'Enemy Vessels
            binWriter.Write(mA8h) : offset += 2   'Sign Marker
            binWriter.Write(mAAh) : offset += 2
            binWriter.Write(mMoveCount) : offset += 2

            SetChanged(False)
        Finally
            If binWriter IsNot Nothing Then binWriter.Close() : binWriter = Nothing
        End Try
    End Sub
#End Region
End Class
