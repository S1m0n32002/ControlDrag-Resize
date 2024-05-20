Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Security.Policy
Public Class ControlDragAndResize
    Private Structure CtrlLimits
        Friend Left As Integer
        Friend Top As Integer
        Friend Right As Integer
        Friend Bottom As Integer
    End Structure

    Private Enum EnmResizing
        None = -1
        Left
        TopLeft
        Top
        TopRight
        Right
        BottomRight
        Bottom
        BottomLeft
    End Enum

    Public Shared Property Margin As Integer = 5

    Private Shared MousePressed As Boolean = False
    Private Shared Resizing As EnmResizing = EnmResizing.None
    Private Shared Moving As Boolean = False
    Private Shared MousePressedPosition As Point = Nothing

    Private Shared ResizMem As New Dictionary(Of Control, Boolean)

    ''' <summary>
    ''' Enables drag and resize for the specified control.
    ''' </summary>
    ''' <param name="DragOnly">Spacifies if the control can only be dragged.</param>
    Public Shared Sub EnableDragAndResize(Control As Control, Optional DragOnly As Boolean = False)
        AddHandler Control.MouseMove, AddressOf DragAndResizeHandler
        AddHandler Control.MouseUp, AddressOf DragAndResizeReleaseHandler
        ResizMem.Add(Control, Not DragOnly)
    End Sub

    ''' <summary>
    ''' Disables both drag and resize for the specified control.
    ''' </summary>
    Public Shared Sub DisableDragAndResize(Control As Control)
        RemoveHandler Control.MouseMove, AddressOf DragAndResizeHandler
        RemoveHandler Control.MouseUp, AddressOf DragAndResizeReleaseHandler
        ResizMem.Remove(Control)
    End Sub

    Private Shared Function CheckMinMax(Size As Size, MinSize As Size, MaxSize As Size, OldLocation As Point, ByRef NewLocation As Point) As Size

        If MinSize.Height > 0 AndAlso Size.Height <= MinSize.Height Then
            Size.Height = MinSize.Height
            If OldLocation <> Nothing AndAlso NewLocation <> Nothing Then NewLocation.Y = OldLocation.Y
        End If
        If MaxSize.Height > 0 AndAlso Size.Height >= MaxSize.Height Then
            Size.Height = MaxSize.Height
            If OldLocation <> Nothing AndAlso NewLocation <> Nothing Then NewLocation.Y = OldLocation.Y
        End If
        If MinSize.Width > 0 AndAlso Size.Width <= MinSize.Width Then
            Size.Width = MinSize.Width
            If OldLocation <> Nothing AndAlso NewLocation <> Nothing Then NewLocation.X = OldLocation.X
        End If
        If MaxSize.Width > 0 AndAlso Size.Width >= MaxSize.Width Then
            Size.Width = MaxSize.Width
            If OldLocation <> Nothing AndAlso NewLocation <> Nothing Then NewLocation.X = OldLocation.X
        End If

        Return Size
    End Function

    Private Shared Sub DragAndResizeHandler(sender As Object, e As MouseEventArgs)
        Dim senderControl As Control = TryCast(sender, Control)

        If senderControl Is Nothing Then Throw New ArgumentException($"Argument type ""{sender.GetType}"" is not compatible.", "sender") : Exit Sub

        Dim SenderParentLimits As New CtrlLimits
        SenderParentLimits.Left = senderControl.Parent.PointToScreen(Point.Empty).X
        SenderParentLimits.Top = senderControl.Parent.PointToScreen(Point.Empty).Y
        SenderParentLimits.Right = SenderParentLimits.Left + senderControl.Parent.Width
        SenderParentLimits.Bottom = SenderParentLimits.Top + senderControl.Parent.Height

        Dim SenderLimits As New CtrlLimits
        SenderLimits.Left = senderControl.PointToScreen(Point.Empty).X
        SenderLimits.Top = senderControl.PointToScreen(Point.Empty).Y
        SenderLimits.Right = SenderLimits.Left + senderControl.Width
        SenderLimits.Bottom = SenderLimits.Top + senderControl.Height

        If e.Button = MouseButtons.Left Then
            If MousePressed = False Then
                MousePressed = True
                MousePressedPosition = New Point(e.X, e.Y)
            End If

#If DEBUG Then
            Dim R As New Pen(Brushes.Red)
            R.DashStyle = Drawing2D.DashStyle.Dash
            Dim PathParent As New Drawing2D.GraphicsPath
            PathParent.AddRectangle(New Rectangle(senderControl.Parent.PointToClient(New Point(SenderParentLimits.Left, SenderParentLimits.Top)).X, senderControl.Parent.PointToClient(New Point(SenderParentLimits.Left, SenderParentLimits.Top)).Y, senderControl.Parent.Width, senderControl.Parent.Height))
            Dim GraphParent As Graphics = TryCast(sender.parent, Control).CreateGraphics()
            GraphParent.DrawPath(R, PathParent)

            Dim B As New Pen(Brushes.Blue)
            B.DashStyle = Drawing2D.DashStyle.Dot
            Dim PathSender As New Drawing2D.GraphicsPath
            PathSender.AddRectangle(New Rectangle(senderControl.PointToClient(New Point(SenderLimits.Left, SenderLimits.Top)).X, senderControl.PointToClient(New Point(SenderLimits.Left, SenderLimits.Top)).Y, senderControl.Width, senderControl.Height))
            Dim GraphSender As Graphics = senderControl.CreateGraphics()
            GraphSender.DrawPath(B, PathSender)

            Debug.Write($"SENDER -> left: {SenderParentLimits.Left} | top: {SenderParentLimits.Top} | right: {SenderParentLimits.Right} | bottom: {SenderParentLimits.Bottom} |" & vbCrLf)
            Debug.Write($"PARENT -> left: {SenderLimits.Left} | top: {SenderLimits.Top} | right: {SenderLimits.Right} | bottom: {SenderLimits.Bottom} |" & vbCrLf)
            Debug.Write($"MOUSE  -> X: {Cursor.Position.X} | Y: {Cursor.Position.Y} |" & vbCrLf)
#End If

            Dim BorderReached As Boolean = False
            If Cursor.Position.X < SenderParentLimits.Left Then Cursor.Position = New Point(SenderParentLimits.Left, Cursor.Position.Y) : BorderReached = True
            If Cursor.Position.X > SenderParentLimits.Right Then Cursor.Position = New Point(SenderParentLimits.Right, Cursor.Position.Y) : BorderReached = True
            If Cursor.Position.Y < SenderParentLimits.Top Then Cursor.Position = New Point(Cursor.Position.X, SenderParentLimits.Top) : BorderReached = True
            If Cursor.Position.Y > SenderParentLimits.Bottom Then Cursor.Position = New Point(Cursor.Position.X, SenderParentLimits.Bottom) : BorderReached = True
            If BorderReached Then Exit Sub

            Select Case Resizing
                Case EnmResizing.Left
                    If Cursor.Current <> Cursors.SizeWE Then Cursor.Current = Cursors.SizeWE

                    Dim newSize As New Size(senderControl.Size.Width - (e.X - MousePressedPosition.X), senderControl.Size.Height)
                    Dim newLocation As New Point(senderControl.Location.X + e.X - MousePressedPosition.X, senderControl.Location.Y)

                    senderControl.Size = CheckMinMax(newSize, senderControl.MinimumSize, senderControl.MaximumSize, senderControl.Location, newLocation)
                    senderControl.Location = newLocation

                Case EnmResizing.TopLeft
                    If Cursor.Current <> Cursors.SizeNWSE Then Cursor.Current = Cursors.SizeNWSE

                    Dim newSize As New Size(senderControl.Size.Width - (e.X - MousePressedPosition.X), senderControl.Size.Height - (e.Y - MousePressedPosition.Y))
                    Dim newLocation As New Point(senderControl.Location.X + e.X - MousePressedPosition.X, senderControl.Location.Y + e.Y - MousePressedPosition.Y)

                    senderControl.Size = CheckMinMax(newSize, senderControl.MinimumSize, senderControl.MaximumSize, senderControl.Location, newLocation)
                    senderControl.Location = newLocation

                Case EnmResizing.Top
                    If Cursor.Current <> Cursors.SizeNS Then Cursor.Current = Cursors.SizeNS

                    Dim newSize As New Size(senderControl.Size.Width, senderControl.Size.Height - (e.Y - MousePressedPosition.Y))
                    Dim newLocation As New Point(senderControl.Location.X, senderControl.Location.Y + e.Y - MousePressedPosition.Y)

                    senderControl.Size = CheckMinMax(newSize, senderControl.MinimumSize, senderControl.MaximumSize, senderControl.Location, newLocation)
                    senderControl.Location = newLocation

                Case EnmResizing.TopRight
                    If Cursor.Current <> Cursors.SizeNESW Then Cursor.Current = Cursors.SizeNESW

                    Dim newSize As New Size(e.X, senderControl.Size.Height - (e.Y - MousePressedPosition.Y))
                    Dim newLocation As New Point(senderControl.Location.X, senderControl.Location.Y + e.Y - MousePressedPosition.Y)

                    senderControl.Size = CheckMinMax(newSize, senderControl.MinimumSize, senderControl.MaximumSize, senderControl.Location, newLocation)
                    senderControl.Location = newLocation

                Case EnmResizing.Right
                    If Cursor.Current <> Cursors.SizeWE Then Cursor.Current = Cursors.SizeWE

                    senderControl.Size = New Size(e.X, senderControl.Size.Height)

                Case EnmResizing.BottomRight
                    If Cursor.Current <> Cursors.SizeNWSE Then Cursor.Current = Cursors.SizeNWSE

                    senderControl.Size = New Size(e.X, e.Y)

                Case EnmResizing.Bottom
                    If Cursor.Current <> Cursors.SizeNS Then Cursor.Current = Cursors.SizeNS

                    senderControl.Size = New Size(senderControl.Size.Width, e.Y)

                Case EnmResizing.BottomLeft
                    If Cursor.Current <> Cursors.SizeNESW Then Cursor.Current = Cursors.SizeNESW

                    Dim newSize As New Size(senderControl.Size.Width - (e.X - MousePressedPosition.X), e.Y)
                    Dim newLocation As New Point(senderControl.Location.X + (e.X - MousePressedPosition.X), senderControl.Location.Y)

                    senderControl.Size = CheckMinMax(newSize, senderControl.MinimumSize, senderControl.MaximumSize, senderControl.Location, newLocation)
                    senderControl.Location = newLocation

                Case Else
                    If Not ResizMem(senderControl) Then Exit Select
                    If Moving Then Exit Select

                    If (Cursor.Position.X <= SenderLimits.Left + Margin And Cursor.Position.Y <= SenderLimits.Top + Margin) Then Resizing = EnmResizing.TopLeft : Exit Sub
                    If (Cursor.Position.X >= SenderLimits.Right - Margin And Cursor.Position.Y <= SenderLimits.Top + Margin) Then Resizing = EnmResizing.TopRight : Exit Sub
                    If (Cursor.Position.X >= SenderLimits.Right - Margin And Cursor.Position.Y >= SenderLimits.Bottom - Margin) Then Resizing = EnmResizing.BottomRight : Exit Sub
                    If (Cursor.Position.X <= SenderLimits.Left + Margin And Cursor.Position.Y >= SenderLimits.Bottom - Margin) Then Resizing = EnmResizing.BottomLeft : Exit Sub
                    If Cursor.Position.X <= SenderLimits.Left + Margin Then Resizing = EnmResizing.Left : Exit Sub
                    If Cursor.Position.Y <= SenderLimits.Top + Margin Then Resizing = EnmResizing.Top : Exit Sub
                    If Cursor.Position.X >= SenderLimits.Right - Margin Then Resizing = EnmResizing.Right : Exit Sub
                    If Cursor.Position.Y >= SenderLimits.Bottom - Margin Then Resizing = EnmResizing.Bottom : Exit Sub

            End Select

            If Resizing = EnmResizing.None Then
                Cursor.Current = Cursors.SizeAll

                Moving = True

                senderControl.Location = New Point(senderControl.Location.X + e.X - MousePressedPosition.X, senderControl.Location.Y + e.Y - MousePressedPosition.Y)
            End If
        Else
            If ResizMem(senderControl) Then
                If (Cursor.Position.X <= SenderLimits.Left + Margin And Cursor.Position.Y <= SenderLimits.Top + Margin) Or (Cursor.Position.X >= SenderLimits.Right - Margin And Cursor.Position.Y >= SenderLimits.Bottom - Margin) Then If Cursor.Current <> Cursors.SizeNWSE Then Cursor.Current = Cursors.SizeNWSE : Exit Sub
                If (Cursor.Position.X >= SenderLimits.Right - Margin And Cursor.Position.Y <= SenderLimits.Top + Margin) Or (Cursor.Position.X <= SenderLimits.Left + Margin And Cursor.Position.Y >= SenderLimits.Bottom - Margin) Then If Cursor.Current <> Cursors.SizeNESW Then Cursor.Current = Cursors.SizeNESW : Exit Sub
                If Cursor.Position.X <= SenderLimits.Left + Margin Or Cursor.Position.X >= SenderLimits.Right - Margin Then If Cursor.Current <> Cursors.SizeWE Then Cursor.Current = Cursors.SizeWE : Exit Sub
                If Cursor.Position.Y <= SenderLimits.Top + Margin Or Cursor.Position.Y >= SenderLimits.Bottom - Margin Then If Cursor.Current <> Cursors.SizeNS Then Cursor.Current = Cursors.SizeNS : Exit Sub
            End If
        End If
    End Sub

    Private Shared Sub DragAndResizeReleaseHandler(ByVal sender As Object, ByVal e As MouseEventArgs)
        MousePressed = False
        Resizing = EnmResizing.None
        Moving = False
        Cursor.Current = Cursors.Default

#If DEBUG Then
        sender.parent.refresh()
        Debug.WriteLine("Mouse rilasciato")
#End If
    End Sub

End Class
