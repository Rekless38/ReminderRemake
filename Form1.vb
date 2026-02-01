Imports System.Windows.Forms

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If COMSpecies.Items.Count = 0 Then
            COMSpecies.Items.Add("Canine")
            COMSpecies.Items.Add("Feline")
        End If
        If COMInc.Items.Count = 0 Then
            COMInc.Items.Add("week(s)")
            COMInc.Items.Add("month(s)")
            COMInc.Items.Add("year(s)")
        End If
        NUMAge.Minimum = 0
        NUMAge.Maximum = 30
    End Sub

    Private Sub BTNCalc_Click(sender As Object, e As EventArgs) Handles BTNCalc.Click
        Dim pet As String = TXBName.Text
        Dim ageInput As Integer = CInt(NUMAge.Value)
        Dim increment As String = COMInc.Text
        Dim ageDisplay As String = ageInput.ToString() & " " & increment

        ' Convert to weeks for schedule logic
        Dim ageInWeeks As Integer
        Select Case increment
            Case "week(s)"
                ageInWeeks = ageInput
            Case "month(s)"
                ageInWeeks = ageInput * 4
            Case "year(s)"
                ageInWeeks = ageInput * 52
            Case Else
                ageInWeeks = ageInput
        End Select

        Dim k9distSchedule, feldistSchedule, bordSchedule, leukSchedule, rabiesSchedule, lymeSchedule, leptoSchedule As String

        Dim xfvrcp As Boolean = CXBFvrcp.Checked
        Dim xleuk As Boolean = CXBLeuk.Checked
        Dim xferabies As Boolean = CXBRabFe.Checked
        Dim xcarabies As Boolean = CXBRabCa.Checked
        Dim xda2pp As Boolean = CXBDa2PP.Checked
        Dim xbord As Boolean = CXBBord.Checked
        Dim xlepto As Boolean = CXBLepto.Checked
        Dim xlyme As Boolean = CXBLyme.Checked

        LBLResultsName.Visible = True
        LBLResultsName.Text = "The age of " & pet & " is: " & ageDisplay

        If COMSpecies.Text = "Feline" Then
            ' FVRCP - always show schedule
            Select Case ageInWeeks
                Case 0 To 5
                    feldistSchedule = "Not old enough for FVRCP vaccine"
                Case 6 To 8
                    feldistSchedule = "FVRCP Initial (6-8 weeks)"
                Case 9 To 12
                    feldistSchedule = "FVRCP Second (9-12 weeks)"
                Case 13 To 16
                    feldistSchedule = "FVRCP Third (13+ weeks)"
                Case 17 To 52
                    feldistSchedule = "FVRCP Annual (1yr)"
                Case Else ' >1 year
                    feldistSchedule = If(xfvrcp, "FVRCP Tri-Annual (3yr)", "FVRCP Annual (1yr)")
            End Select
            LBLResultsDist.Visible = True
            LBLResultsDist.Text = feldistSchedule

            ' Leukemia - always show schedule (uses Bord label)
            Select Case ageInWeeks
                Case 0 To 8
                    leukSchedule = "Not old enough for Leukemia vaccine"
                Case 9 To 12
                    leukSchedule = "Feline Leukemia Initial (9-12 weeks)"
                Case 13 To 52
                    leukSchedule = "Feline Leukemia Booster (13+ weeks)"
                Case Else ' >1 year
                    leukSchedule = If(xleuk, "Feline Leukemia Annual (1yr)", "Feline Leukemia Annual (1yr)")
            End Select
            LBLResultsBord.Visible = True
            LBLResultsBord.Text = leukSchedule

            ' Rabies Feline - always show
            Select Case ageInWeeks
                Case 0 To 11
                    rabiesSchedule = "Not old enough for Rabies vaccine"
                Case 12 To 52
                    rabiesSchedule = "Feline Rabies Annual (1yr)"
                Case Else
                    rabiesSchedule = If(xferabies, "Feline Rabies Tri-Annual (3yr)", "Feline Rabies Annual (1yr)")
            End Select
            LBLResultsRabie.Visible = True
            LBLResultsRabie.Text = rabiesSchedule

            LBLResultsLyme.Visible = False
            LBLResultsLepto.Visible = False

        ElseIf COMSpecies.Text = "Canine" Then
            ' DA2PP - always show
            Select Case ageInWeeks
                Case 0 To 5
                    k9distSchedule = "Not old enough for Da2PP vaccine"
                Case 6 To 8
                    k9distSchedule = "Da2PP Initial (6-8 weeks)"
                Case 9 To 12
                    k9distSchedule = "Da2PP Second (9-12 weeks)"
                Case 13 To 16
                    k9distSchedule = "Da2PP Third (13+ weeks)"
                Case 17 To 52
                    k9distSchedule = "Da2PP Annual (1yr)"
                Case Else ' >1 year
                    k9distSchedule = If(xda2pp, "Da2PP TriAnnual (3yr)", "Da2PP Annual (1yr)")
            End Select
            LBLResultsDist.Visible = True
            LBLResultsDist.Text = k9distSchedule

            ' Bordetella - always show appropriate message
            Select Case ageInWeeks
                Case 0 To 5
                    bordSchedule = "Not old enough for Bordetella vaccine"
                Case Else
                    bordSchedule = If(xbord, "Canine Bordetella Annual", "Canine Bordetella Annual")
            End Select
            LBLResultsBord.Visible = True
            LBLResultsBord.Text = bordSchedule

            ' Rabies Canine - always show
            Select Case ageInWeeks
                Case 0 To 11
                    rabiesSchedule = "Not old enough for Rabies"
                Case 12 To 52
                    rabiesSchedule = "Canine Rabies Annual (1yr)"
                Case Else
                    rabiesSchedule = If(xcarabies, "Canine Rabies TriAnnual (3yr)", "Canine Rabies Annual (1yr)")
            End Select
            LBLResultsRabie.Visible = True
            LBLResultsRabie.Text = rabiesSchedule

            ' Leptospirosis - always show (shown in Lyme label per your original)
            Select Case ageInWeeks
                Case 0 To 8
                    leptoSchedule = "Not old enough for Leptospirosis vaccine"
                Case 9 To 12
                    leptoSchedule = "Leptospirosis Initial (9-12 weeks)"
                Case 13 To 52
                    leptoSchedule = "Leptospirosis Booster (13+ weeks)"
                Case Else
                    leptoSchedule = If(xlepto, "Leptospirosis Annual (1yr)", "Leptospirosis Annual (1yr)")
            End Select
            LBLResultsLyme.Visible = True
            LBLResultsLyme.Text = leptoSchedule

            ' Lyme - always show (shown in Lepto label per your original)
            Select Case ageInWeeks
                Case 0 To 8
                    lymeSchedule = "Not old enough for Lyme vaccine"
                Case 9 To 12
                    lymeSchedule = "Canine Lyme Initial (9-12 weeks)"
                Case 13 To 52
                    lymeSchedule = "Canine Lyme Booster (13+ weeks)"
                Case Else
                    lymeSchedule = If(xlyme, "Canine Lyme Annual (1yr)", "Canine Lyme Annual (1yr)")
            End Select
            LBLResultsLepto.Visible = True
            LBLResultsLepto.Text = lymeSchedule
        End If
    End Sub

    Private Sub BTNClear_Click(sender As Object, e As EventArgs) Handles BTNClear.Click
        TXBName.Text = ""
        COMSpecies.Text = ""
        COMInc.Text = ""
        NUMAge.Value = 0
        LBLResultsDist.Visible = False
        LBLResultsBord.Visible = False
        LBLResultsRabie.Visible = False
        LBLResultsLyme.Visible = False
        LBLResultsLepto.Visible = False
        CXBBord.Checked = False
        CXBDa2PP.Checked = False
        CXBFvrcp.Checked = False
        CXBLepto.Checked = False
        CXBLeuk.Checked = False
        CXBLyme.Checked = False
        CXBRabCa.Checked = False
        CXBRabFe.Checked = False
        If Me.Controls.Contains(CXBUnknown) Then CXBUnknown.Checked = False
        TXBName.Focus()
    End Sub

    Private Sub BTNExit_Click(sender As Object, e As EventArgs) Handles BTNExit.Click
        Me.Close()
    End Sub

End Class
