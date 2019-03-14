Imports System.ComponentModel

Public Class frmMain

    Private Sub tcMain_DrawItem(sender As Object, e As DrawItemEventArgs) Handles tcMain.DrawItem
        Dim g As Graphics = e.Graphics
        Dim _TextBrush As Brush

        ' Get the item from the collection.
        Dim _TabPage As TabPage = tcMain.TabPages(e.Index)

        ' Get the real bounds for the tab rectangle.
        Dim _TabBounds As Rectangle = tcMain.GetTabRect(e.Index)

        If (e.State = DrawItemState.Selected) Then
            ' Draw a different background color, and don't paint a focus rectangle.
            _TextBrush = New SolidBrush(Color.Orange)
            g.FillRectangle(Brushes.DimGray, e.Bounds)
        Else
            _TextBrush = New System.Drawing.SolidBrush(e.ForeColor)
            e.DrawBackground()
        End If

        ' Use our own font.
        Dim _TabFont As New Font("Century Gothic", 11.0, FontStyle.Regular, GraphicsUnit.Pixel)

        ' Draw string. Center the text.
        Dim _StringFlags As New StringFormat()
        _StringFlags.Alignment = StringAlignment.Center
        _StringFlags.LineAlignment = StringAlignment.Center
        g.DrawString(_TabPage.Text, _TabFont, _TextBrush, _TabBounds, New StringFormat(_StringFlags))
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        lblTitle.Text = "Inflation Calculator v" & System.Windows.Forms.Application.ProductVersion

        loadYear()
    End Sub

    Private Sub loadYear()
        Dim inflationYears As Dictionary(Of Integer, Decimal) = loadInflationYear()

        For Each inflationYear In inflationYears
            Dim oItem As New clsKeyValuePair(inflationYear.Key, inflationYear.Value)

            cmbFromYear.Items.Add(oItem)
            cmbToYear.Items.Add(oItem)
        Next

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Clear()
    End Sub

    Public Function proccessApi() As Boolean
        Dim totalValue As Decimal = txtDesiredvalue.Text
        Dim inflationyear = loadInflationYear()
        Dim datestart = CInt(cmbFromYear.Text)
        Dim dateend = CInt(cmbToYear.Text)

        For current As Integer = datestart To dateend
            totalValue = ((inflationyear.Item(current) / 100) * totalValue) + totalValue
        Next

        txtResultValue.Text = totalValue.ToString("0.00")
        txtResultPercentage.Text = (((totalValue - txtDesiredvalue.Text) / txtDesiredvalue.Text) * 100).ToString("0.00") & "%"
        Return True
    End Function

    Private Sub Clear()
        lblDesiredValue.ForeColor = Color.Silver
        lblFromYear.ForeColor = Color.Silver
        lblToYear.ForeColor = Color.Silver

        txtDesiredvalue.Text = "0.00"
        cmbFromYear.SelectedIndex = -1
        cmbToYear.SelectedIndex = -1
        txtResultValue.Text = "0.00"
        txtResultPercentage.Text = ""
    End Sub

    Private Sub txtDesiredvalue_Validated(sender As Object, e As EventArgs) Handles txtDesiredvalue.Validated
        If Not IsNumeric(sender.Text) Then
            MsgBox("Decimal Value ONLY", MsgBoxStyle.ApplicationModal, "Warning")

            txtDesiredvalue.Text = "0.00"
            Return
        End If

        sender.Text = CDec(sender.Text).ToString("0.00")
    End Sub

    Private Sub btnCompute_Click(sender As Object, e As EventArgs) Handles btnCompute.Click

        If Val(txtDesiredvalue.Text) = 0 Or String.IsNullOrEmpty(cmbFromYear.Text) Or String.IsNullOrEmpty(cmbToYear.Text) Then
            MsgBox("Kindly fill in the required fields", MsgBoxStyle.ApplicationModal, "System Warning")
            lblDesiredValue.ForeColor = Color.Orange
            lblFromYear.ForeColor = Color.Orange
            lblToYear.ForeColor = Color.Orange
            Return
        End If

        If CInt(cmbFromYear.Text) >= CInt(cmbToYear.Text) Then
            MsgBox("[ From Year ] must not be ahead of [ To Year ]", MsgBoxStyle.ApplicationModal, "System Warning")
            lblFromYear.ForeColor = Color.Orange
            lblToYear.ForeColor = Color.Orange
            Return
        End If

        proccessApi()



    End Sub

End Class