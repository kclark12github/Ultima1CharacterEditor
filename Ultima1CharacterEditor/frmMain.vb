Imports System.ComponentModel

Public Class frmMain
    Private mPlayer As Player = Nothing
    Private mPlayer1 As Player = Nothing
    Private mPlayer2 As Player = Nothing
    Private mPlayer3 As Player = Nothing
    Private mPlayer4 As Player = Nothing
    Private Sub cbClass_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbClass.SelectedIndexChanged
        mPlayer.Class = cbClass.SelectedIndex
    End Sub
    Private Sub cbRace_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbRace.SelectedIndexChanged
        mPlayer.Race = cbRace.SelectedIndex
    End Sub
    Private Sub cbReadyArmor_SelectedIndexChanged(sender As Object, e As EventArgs)
        mPlayer.ReadyArmor = cbReadyArmor.SelectedIndex
    End Sub
    Private Sub cbReadySpell_SelectedIndexChanged(sender As Object, e As EventArgs)
        mPlayer.ReadySpell = cbReadySpell.SelectedIndex
    End Sub
    Private Sub cbReadyWeapon_SelectedIndexChanged(sender As Object, e As EventArgs)
        mPlayer.ReadyWeapon = cbReadyWeapon.SelectedIndex
    End Sub
    Private Sub cbSex_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbSex.SelectedIndexChanged
        mPlayer.Sex = cbSex.SelectedIndex
    End Sub
    Private Sub cbPlayers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbPlayers.SelectedIndexChanged
        If cbPlayers.Items(cbPlayers.SelectedIndex).StartsWith("Player") Then
            MessageBox.Show(Me, "No Player Data File Exists.", "")
        Else
            If mPlayer IsNot Nothing AndAlso mPlayer.Changed Then
                If MessageBox.Show(Me, "OK to Save Changes?", "Save?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then cmdSave_Click(sender, e) Else mPlayer1.Read()
            End If
            Select Case cbPlayers.SelectedIndex
                Case 0 : mPlayer = mPlayer1
                Case 1 : mPlayer = mPlayer2
                Case 2 : mPlayer = mPlayer3
                Case 3 : mPlayer = mPlayer4
            End Select

            lblFileName.Text = mPlayer.FileName

            cbClass.SelectedIndex = mPlayer.Class
            cbRace.SelectedIndex = mPlayer.Race
            cbSex.SelectedIndex = mPlayer.Sex
            cbReadyArmor.SelectedIndex = mPlayer.ReadyArmor
            cbReadySpell.SelectedIndex = mPlayer.ReadySpell
            cbReadyWeapon.SelectedIndex = mPlayer.ReadyWeapon
            cbTransport.SelectedIndex = mPlayer.Transport

            With nudHits : .Minimum = 0 : .Maximum = mPlayer.maxValue : .Value = mPlayer.Hits : End With
            With nudExperience : .Minimum = 0 : .Maximum = mPlayer.maxValue : .Value = mPlayer.Experience : End With
            With nudCoin : .Minimum = 0 : .Maximum = mPlayer.maxValue : .Value = mPlayer.Coin : End With
            With nudFood : .Minimum = 0 : .Maximum = mPlayer.maxValue : .Value = mPlayer.Food : End With

            With nudStrength : .Minimum = 0 : .Maximum = mPlayer.maxAttribute : .Value = mPlayer.Strength : End With
            With nudAgility : .Minimum = 0 : .Maximum = mPlayer.maxAttribute : .Value = mPlayer.Agility : End With
            With nudStamina : .Minimum = 0 : .Maximum = mPlayer.maxAttribute : .Value = mPlayer.Stamina : End With
            With nudCharisma : .Minimum = 0 : .Maximum = mPlayer.maxAttribute : .Value = mPlayer.Charisma : End With
            With nudWisdom : .Minimum = 0 : .Maximum = mPlayer.maxAttribute : .Value = mPlayer.Wisdom : End With
            With nudIntelligence : .Minimum = 0 : .Maximum = mPlayer.maxAttribute : .Value = mPlayer.Intelligence : End With

            With nudBlueGem : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.BlueGem : End With
            With nudGreenGem : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.GreenGem : End With
            With nudRedGem : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.RedGem : End With
            With nudWhiteGem : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WhiteGem : End With

            With nudLeather : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.ArmorLeather : End With
            With nudChainMail : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.ArmorChainMail : End With
            With nudPlateMail : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.ArmorPlateMail : End With
            With nudReflectSuit : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.ArmorReflectSuit : End With
            With nudVacuumSuit : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.ArmorVacuumSuit : End With

            With nudDagger : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponDagger : End With
            With nudMace : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponMace : End With
            With nudAxe : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponAxe : End With
            With nudRopeAndSpikes : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponRopeAndSpikes : End With
            With nudSword : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponSword : End With
            With nudGreatSword : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponGreatSword : End With
            With nudBowAndArrows : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponBowAndArrows : End With
            With nudAmulet : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponAmulet : End With
            With nudWand : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponWand : End With
            With nudStaff : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponStaff : End With
            With nudTriangle : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponTriangle : End With
            With nudPistol : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponPistol : End With
            With nudLightSword : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponLightSword : End With
            With nudPhazor : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponPhazor : End With
            With nudBlaster : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.WeaponBlaster : End With

            With nudOpen : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellOpen : End With
            With nudUnlock : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellUnlock : End With
            With nudMagicMissile : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellMagicMissile : End With
            With nudSteal : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellSteal : End With
            With nudLadderDown : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellLadderDown : End With
            With nudLadderUp : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellLadderUp : End With
            With nudBlink : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellBlink : End With
            With nudCreate : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellCreate : End With
            With nudDestroy : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellDestroy : End With
            With nudKill : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.SpellKill : End With

            With nudHorse : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportHorse : End With
            With nudCart : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportCart : End With
            With nudRaft : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportRaft : End With
            With nudFrigate : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportFrigate : End With
            With nudAircar : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportAircar : End With
            With nudShuttle : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportShuttle : End With
            With nudTimeMachine : .Minimum = 0 : .Maximum = mPlayer.maxItems : .Value = mPlayer.TransportTimeMachine : End With
        End If
        cmdSave.Enabled = False
    End Sub
    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        If mPlayer IsNot Nothing AndAlso mPlayer.Changed Then
            If MessageBox.Show(Me, "OK to Save Changes?", "Save?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then cmdSave_Click(sender, e)
        End If
        Me.Close()
    End Sub
    Private Sub cmdMax_Click(sender As Object, e As EventArgs)
        Dim ctl As Button = sender
        Select Case ctl.Name
            Case "cmdMaxCoin" : nudCoin.Value = nudCoin.Maximum
            Case "cmdMaxExperience" : nudExperience.Value = nudExperience.Maximum
            Case "cmdMaxFood" : nudFood.Value = nudFood.Maximum
            Case "cmdMaxHits" : nudHits.Value = nudHits.Maximum

            Case "cmdMaxAgility" : nudAgility.Value = nudAgility.Maximum
            Case "cmdMaxCharisma" : nudCharisma.Value = nudCharisma.Maximum
            Case "cmdMaxIntelligence" : nudIntelligence.Value = nudIntelligence.Maximum
            Case "cmdMaxStamina" : nudStamina.Value = nudStamina.Maximum
            Case "cmdMaxStrength" : nudStrength.Value = nudStrength.Maximum
            Case "cmdMaxWisdom" : nudWisdom.Value = nudWisdom.Maximum

            Case "cmdMaxBlueGem" : nudBlueGem.Value = nudBlueGem.Maximum
            Case "cmdMaxGreenGem" : nudGreenGem.Value = nudGreenGem.Maximum
            Case "cmdMaxRedGem" : nudRedGem.Value = nudRedGem.Maximum
            Case "cmdMaxWhiteGem" : nudWhiteGem.Value = nudWhiteGem.Maximum

            Case "cmdMaxLeather" : nudLeather.Value = nudLeather.Maximum
            Case "cmdMaxChainMail" : nudChainMail.Value = nudChainMail.Maximum
            Case "cmdMaxPlateMail" : nudPlateMail.Value = nudPlateMail.Maximum
            Case "cmdMaxVacuumSuit" : nudVacuumSuit.Value = nudVacuumSuit.Maximum
            Case "cmdMaxReflectSuit" : nudReflectSuit.Value = nudReflectSuit.Maximum

            Case "cmdMaxDagger" : nudDagger.Value = nudDagger.Maximum
            Case "cmdMaxMace" : nudMace.Value = nudMace.Maximum
            Case "cmdMaxAxe" : nudAxe.Value = nudAxe.Maximum
            Case "cmdMaxRopeAndSpikes" : nudRopeAndSpikes.Value = nudRopeAndSpikes.Maximum
            Case "cmdMaxSword" : nudSword.Value = nudSword.Maximum
            Case "cmdMaxGreatSword" : nudGreatSword.Value = nudGreatSword.Maximum
            Case "cmdMaxBowAndArrows" : nudBowAndArrows.Value = nudBowAndArrows.Maximum
            Case "cmdMaxAmulet" : nudAmulet.Value = nudAmulet.Maximum
            Case "cmdMaxWand" : nudWand.Value = nudWand.Maximum
            Case "cmdMaxStaff" : nudStaff.Value = nudStaff.Maximum
            Case "cmdMaxTriangle" : nudTriangle.Value = nudTriangle.Maximum
            Case "cmdMaxPistol" : nudPistol.Value = nudPistol.Maximum
            Case "cmdMaxLightSword" : nudLightSword.Value = nudLightSword.Maximum
            Case "cmdMaxPhazor" : nudPhazor.Value = nudPhazor.Maximum
            Case "cmdMaxBlaster" : nudBlaster.Value = nudBlaster.Maximum

            Case "cmdMaxOpen" : nudOpen.Value = nudOpen.Maximum
            Case "cmdMaxUnlock" : nudUnlock.Value = nudUnlock.Maximum
            Case "cmdMaxMagicMissile" : nudMagicMissile.Value = nudMagicMissile.Maximum
            Case "cmdMaxSteal" : nudSteal.Value = nudSteal.Maximum
            Case "cmdMaxLadderDown" : nudLadderDown.Value = nudLadderDown.Maximum
            Case "cmdMaxLadderUp" : nudLadderUp.Value = nudLadderUp.Maximum
            Case "cmdMaxBlink" : nudBlink.Value = nudBlink.Maximum
            Case "cmdMaxCreate" : nudCreate.Value = nudCreate.Maximum
            Case "cmdMaxDestroy" : nudDestroy.Value = nudDestroy.Maximum
            Case "cmdMaxKill" : nudKill.Value = nudKill.Maximum

            Case "cmdMaxHorse" : nudHorse.Value = nudHorse.Maximum
            Case "cmdMaxCart" : nudCart.Value = nudCart.Maximum
            Case "cmdMaxRaft" : nudRaft.Value = nudRaft.Maximum
            Case "cmdMaxFrigate" : nudFrigate.Value = nudFrigate.Maximum
            Case "cmdMaxAircar" : nudAircar.Value = nudAircar.Maximum
            Case "cmdMaxShuttle" : nudShuttle.Value = nudShuttle.Maximum
            Case "cmdMaxTimeMachine" : nudTimeMachine.Value = nudTimeMachine.Maximum
        End Select
    End Sub
    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Try : If mPlayer.Changed Then mPlayer.Save()
        Catch ex As Exception : MessageBox.Show(Me, ex.ToString, ex.GetType.Name)
        Finally : cmdSave.Enabled = mPlayer.Changed
        End Try
    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            lblFileName.Text = ""
            Try : mPlayer1 = New Player(1) : cbPlayers.Items(0) = mPlayer1.Name : Catch ex As FileNotFoundException : mPlayer1 = Nothing : End Try
            Try : mPlayer2 = New Player(2) : cbPlayers.Items(1) = mPlayer2.Name : Catch ex As FileNotFoundException : mPlayer2 = Nothing : End Try
            Try : mPlayer3 = New Player(3) : cbPlayers.Items(2) = mPlayer3.Name : Catch ex As FileNotFoundException : mPlayer3 = Nothing : End Try
            Try : mPlayer4 = New Player(4) : cbPlayers.Items(3) = mPlayer4.Name : Catch ex As FileNotFoundException : mPlayer4 = Nothing : End Try

            AddHandler nudCoin.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxCoin.Click, AddressOf cmdMax_Click
            AddHandler nudExperience.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxExperience.Click, AddressOf cmdMax_Click
            AddHandler nudFood.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxFood.Click, AddressOf cmdMax_Click
            AddHandler nudHits.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxHits.Click, AddressOf cmdMax_Click

            AddHandler nudAgility.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxAgility.Click, AddressOf cmdMax_Click
            AddHandler nudCharisma.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxCharisma.Click, AddressOf cmdMax_Click
            AddHandler nudIntelligence.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxIntelligence.Click, AddressOf cmdMax_Click
            AddHandler nudStamina.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxStamina.Click, AddressOf cmdMax_Click
            AddHandler nudStrength.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxStrength.Click, AddressOf cmdMax_Click
            AddHandler nudWisdom.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxWisdom.Click, AddressOf cmdMax_Click

            AddHandler nudBlueGem.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxBlueGem.Click, AddressOf cmdMax_Click
            AddHandler nudGreenGem.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxGreenGem.Click, AddressOf cmdMax_Click
            AddHandler nudRedGem.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxRedGem.Click, AddressOf cmdMax_Click
            AddHandler nudWhiteGem.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxWhiteGem.Click, AddressOf cmdMax_Click

            AddHandler nudLeather.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxLeather.Click, AddressOf cmdMax_Click
            AddHandler nudChainMail.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxChainMail.Click, AddressOf cmdMax_Click
            AddHandler nudPlateMail.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxPlateMail.Click, AddressOf cmdMax_Click
            AddHandler nudVacuumSuit.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxVacuumSuit.Click, AddressOf cmdMax_Click
            AddHandler nudReflectSuit.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxReflectSuit.Click, AddressOf cmdMax_Click

            AddHandler nudDagger.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxDagger.Click, AddressOf cmdMax_Click
            AddHandler nudMace.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxMace.Click, AddressOf cmdMax_Click
            AddHandler nudAxe.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxAxe.Click, AddressOf cmdMax_Click
            AddHandler nudRopeAndSpikes.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxRopeAndSpikes.Click, AddressOf cmdMax_Click
            AddHandler nudSword.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxSword.Click, AddressOf cmdMax_Click
            AddHandler nudGreatSword.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxGreatSword.Click, AddressOf cmdMax_Click
            AddHandler nudBowAndArrows.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxBowAndArrows.Click, AddressOf cmdMax_Click
            AddHandler nudAmulet.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxAmulet.Click, AddressOf cmdMax_Click
            AddHandler nudWand.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxWand.Click, AddressOf cmdMax_Click
            AddHandler nudStaff.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxStaff.Click, AddressOf cmdMax_Click
            AddHandler nudTriangle.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxTriangle.Click, AddressOf cmdMax_Click
            AddHandler nudPistol.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxPistol.Click, AddressOf cmdMax_Click
            AddHandler nudLightSword.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxLightSword.Click, AddressOf cmdMax_Click
            AddHandler nudPhazor.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxPhazor.Click, AddressOf cmdMax_Click
            AddHandler nudBlaster.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxBlaster.Click, AddressOf cmdMax_Click

            AddHandler nudOpen.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxOpen.Click, AddressOf cmdMax_Click
            AddHandler nudUnlock.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxUnlock.Click, AddressOf cmdMax_Click
            AddHandler nudMagicMissile.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxMagicMissile.Click, AddressOf cmdMax_Click
            AddHandler nudSteal.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxSteal.Click, AddressOf cmdMax_Click
            AddHandler nudLadderDown.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxLadderDown.Click, AddressOf cmdMax_Click
            AddHandler nudLadderUp.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxLadderUp.Click, AddressOf cmdMax_Click
            AddHandler nudBlink.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxBlink.Click, AddressOf cmdMax_Click
            AddHandler nudCreate.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxCreate.Click, AddressOf cmdMax_Click
            AddHandler nudDestroy.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxDestroy.Click, AddressOf cmdMax_Click
            AddHandler nudKill.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxKill.Click, AddressOf cmdMax_Click

            AddHandler nudHorse.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxHorse.Click, AddressOf cmdMax_Click
            AddHandler nudCart.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxCart.Click, AddressOf cmdMax_Click
            AddHandler nudRaft.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxRaft.Click, AddressOf cmdMax_Click
            AddHandler nudFrigate.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxFrigate.Click, AddressOf cmdMax_Click
            AddHandler nudAircar.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxAircar.Click, AddressOf cmdMax_Click
            AddHandler nudShuttle.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxShuttle.Click, AddressOf cmdMax_Click
            AddHandler nudTimeMachine.ValueChanged, AddressOf nud_ValueChanged : AddHandler cmdMaxTimeMachine.Click, AddressOf cmdMax_Click

            cmdSave.Enabled = False
        Catch ex As Exception : MessageBox.Show(Me, ex.ToString, ex.GetType.Name)
        End Try
    End Sub
    Private Sub nud_ValueChanged(sender As Object, e As EventArgs)
        Dim ctl As NumericUpDown = sender
        Select Case ctl.Name
            Case "nudCoin" : mPlayer.Coin = ctl.Value
            Case "nudExperience" : mPlayer.Experience = ctl.Value
            Case "nudFood" : mPlayer.Food = ctl.Value
            Case "nudHits" : mPlayer.Hits = ctl.Value

            Case "nudAgility" : mPlayer.Agility = ctl.Value
            Case "nudCharisma" : mPlayer.Charisma = ctl.Value
            Case "nudIntelligence" : mPlayer.Intelligence = ctl.Value
            Case "nudStamina" : mPlayer.Stamina = ctl.Value
            Case "nudStrength" : mPlayer.Strength = ctl.Value
            Case "nudWisdom" : mPlayer.Wisdom = ctl.Value

            Case "nudBlueGem" : mPlayer.BlueGem = ctl.Value
            Case "nudGreenGem" : mPlayer.GreenGem = ctl.Value
            Case "nudRedGem" : mPlayer.RedGem = ctl.Value
            Case "nudWhiteGem" : mPlayer.WhiteGem = ctl.Value

            Case "nudLeather" : mPlayer.ArmorLeather = ctl.Value
            Case "nudChainMail" : mPlayer.ArmorChainMail = ctl.Value
            Case "nudPlateMail" : mPlayer.ArmorPlateMail = ctl.Value
            Case "nudVacuumSuit" : mPlayer.ArmorVacuumSuit = ctl.Value
            Case "nudReflectSuit" : mPlayer.ArmorReflectSuit = ctl.Value

            Case "nudDagger" : mPlayer.WeaponDagger = ctl.Value
            Case "nudMace" : mPlayer.WeaponMace = ctl.Value
            Case "nudAxe" : mPlayer.WeaponAxe = ctl.Value
            Case "nudRopeAndSpikes" : mPlayer.WeaponRopeAndSpikes = ctl.Value
            Case "nudSword" : mPlayer.WeaponSword = ctl.Value
            Case "nudGreatSword" : mPlayer.WeaponGreatSword = ctl.Value
            Case "nudBowAndArrows" : mPlayer.WeaponBowAndArrows = ctl.Value
            Case "nudAmulet" : mPlayer.WeaponAmulet = ctl.Value
            Case "nudWand" : mPlayer.WeaponWand = ctl.Value
            Case "nudStaff" : mPlayer.WeaponStaff = ctl.Value
            Case "nudTriangle" : mPlayer.WeaponTriangle = ctl.Value
            Case "nudPistol" : mPlayer.WeaponPistol = ctl.Value
            Case "nudLightSword" : mPlayer.WeaponLightSword = ctl.Value
            Case "nudPhazor" : mPlayer.WeaponPhazor = ctl.Value
            Case "nudBlaster" : mPlayer.WeaponBlaster = ctl.Value

            Case "nudOpen" : mPlayer.SpellOpen = ctl.Value
            Case "nudUnlock" : mPlayer.SpellUnlock = ctl.Value
            Case "nudMagicMissile" : mPlayer.SpellMagicMissile = ctl.Value
            Case "nudSteal" : mPlayer.SpellSteal = ctl.Value
            Case "nudLadderDown" : mPlayer.SpellLadderDown = ctl.Value
            Case "nudLadderUp" : mPlayer.SpellLadderUp = ctl.Value
            Case "nudBlink" : mPlayer.SpellBlink = ctl.Value
            Case "nudCreate" : mPlayer.SpellCreate = ctl.Value
            Case "nudDestroy" : mPlayer.SpellDestroy = ctl.Value
            Case "nudKill" : mPlayer.SpellKill = ctl.Value

            Case "nudHorse" : mPlayer.TransportHorse = ctl.Value
            Case "nudCart" : mPlayer.TransportCart = ctl.Value
            Case "nudRaft" : mPlayer.TransportRaft = ctl.Value
            Case "nudFrigate" : mPlayer.TransportFrigate = ctl.Value
            Case "nudAircar" : mPlayer.TransportAircar = ctl.Value
            Case "nudShuttle" : mPlayer.TransportShuttle = ctl.Value
            Case "nudTimeMachine" : mPlayer.TransportTimeMachine = ctl.Value
        End Select
        cmdSave.Enabled = mPlayer.Changed
    End Sub
End Class
