﻿Public Class ConfigSelect

    Public ConfigList As ArrayList

    Private Sub ConfigSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i = 1 To ConfigList.Count - 1
            ConfigListBox.Items.Add(ConfigList.Item(i))
        Next

        ConfigListBox.SelectedIndex = 0

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Me.Close()

    End Sub

End Class