Public Class Form1
    Private En As Boolean = False
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        En = True
        Utils.ControlDragAndResize.EnableDragAndResize(Panel1, True)
        Utils.ControlDragAndResize.EnableDragAndResize(Panel2)
        Utils.ControlDragAndResize.EnableDragAndResize(Label1)
        Utils.ControlDragAndResize.EnableDragAndResize(Label2, True)
    End Sub

    Private Sub Form1_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick
        If En Then
            En = False
            Utils.ControlDragAndResize.DisableDragAndResize(Panel1)
        Else
            En = True
            Utils.ControlDragAndResize.EnableDragAndResize(Panel1)
        End If
    End Sub
End Class
