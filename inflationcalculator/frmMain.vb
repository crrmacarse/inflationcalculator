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

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()
    End Sub
End Class