Public Class frmMessage

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Me.Close()
        Me.Dispose()

    End Sub


    Private Sub frmMessage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click

        Me.Close()
        Me.Dispose()

    End Sub


    Private Sub lblMessage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMessage.Click

        Me.Close()
        Me.Dispose()

    End Sub


    Private Sub tmrUpdate_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrUpdate.Tick

        Me.Close()
        Me.Dispose()

    End Sub


    Public Sub SetMessage(ByVal TheMessage As String)

        With Me
            .lblMessage.Text = TheMessage
        End With
    End Sub


End Class