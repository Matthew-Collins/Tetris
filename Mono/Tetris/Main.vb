Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Collections
Imports System.Collections.Generic
Imports Microsoft.VisualBasic

Public Class Start
	Public Shared Sub Main
		Dim Dialog As New xMain
		Dialog.WindowState = 2
		Dialog.ShowDialog()
	End Sub
End Class

Public Class xMain
    Inherits Form
    Private Const Title = "Tetris"
    Private Timer As New Timer
    Private Speed As New Timer
    Private Grid As New Grid(Me)
    Public Preview As New Grid(Me)
    Public Sub New()

        Me.Text = Title
        Me.ClientSize = New Size(Me.Grid.Size.Width * 2.2, Me.Grid.Size.Height)
        Me.StartPosition = FormStartPosition.CenterScreen

        With Me.Grid
        	.Size = New Size(Grid.x * Tile.x + 2, Grid.y * Tile.y + 2)
        	.BorderStyle = BorderStyle.FixedSingle
        End With
		Me.Controls.Add(Me.Grid)

		With Me.Preview
			.Size = New Size(5 * Tile.x, 4 * Tile.y)
			.Dock = DockStyle.Right
		End With
        Me.Controls.Add(Me.Preview)

        Me.NewGame()
    End Sub
    Private Sub NewGame()
        Me.Grid.Controls.Clear()
        Me.Grid.Shapes = 0
        Me.Grid.NewShape()
        AddHandler Me.Timer.Tick, AddressOf Timer_Tick
        Me.Timer.Interval = 500
        Me.Timer.Start()
        AddHandler Me.Speed.Tick, AddressOf Speed_Tick
        Me.Speed.Interval = 60000
        Me.Speed.Start()
        Me.Grid.Refresh()
    End Sub
    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        If Me.Grid.MoveShapeDown() Then
            Me.Text = Title & " - Score: " & (Me.Grid.Shapes - 1) * 100
        Else
            Me.Timer.Stop()
            MsgBox("Game Over")
            Me.NewGame()
        End If
    End Sub
    Private Sub Speed_Tick(sender As Object, e As EventArgs)
        Me.Timer.Interval *= 0.95
    End Sub
    Private Sub Main_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Up
                If e.Shift Then
                    Me.Grid.ActiveShape.RotateAntiClockwise()
                Else
                    Me.Grid.ActiveShape.RotateClockwise()
                End If
            Case Keys.Left
                Me.Grid.ActiveShape.MoveLeft()
            Case Keys.Right
                Me.Grid.ActiveShape.MoveRight()
            Case Keys.Down
                Me.Grid.ActiveShape.Drop()
        End Select
    End Sub
    Private Sub Main_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Me.Grid.Left = (Me.ClientSize.Width - Me.Grid.Width) / 2
        Me.Grid.Top = (Me.ClientSize.Height - Me.Grid.Height) / 2
    End Sub
End Class

Public Class Grid
    Inherits Panel
    Public Const x = 10
    Public Const y = 20
    Public ActiveShape As Shape
    Public Colors() As Color = {Color.Red, Color.Green, Color.Magenta, Color.Cyan, Color.Yellow, Color.Gray, Color.Black}
    Public ColorIndex As Integer = -1
    Public Random As New Random
    Public NextShape As Shape.Styles?
    Public Shapes As Integer
    Public Form As xMain
    Public Sub New(Form As xMain)
        Me.Form = Form
    End Sub
    Public Function MoveShapeDown() As Boolean
        If Me.ActiveShape.CanMove(0, 1) Then
            Me.ActiveShape.DoMove(0, 1)
        Else
            Me.RemoveCompletedLines()
            Me.Refresh()
            Me.NewShape()
            If Not Me.ActiveShape.CanMove(0, 0) Then
                Return False
            End If
        End If
        Return True
    End Function
    Public Sub NewShape()
        Me.Shapes += 1

        'Get Color
        Me.ColorIndex += 1
        If Me.ColorIndex = Colors.Length Then
            Me.ColorIndex = 0
        End If
        Dim PreviewColor As Integer = Me.ColorIndex + 1
        If PreviewColor = Colors.Length Then
            PreviewColor = 0
        End If

        'Get Shape
        Dim Style As Shape.Styles
        If Not Me.NextShape.HasValue Then
            Me.NextShape = Me.Random.Next(0, 7)
        End If
        Style = Me.NextShape
        'Style = Shape.Styles.I
        Me.NextShape = Me.Random.Next(0, 7)

        'Add Shape
        Me.ActiveShape = New Shape(Me, Colors(ColorIndex), Style)

        'Add Preview Shape
        Me.Form.Preview.Controls.Clear()
        Dim PreviewShape As New Shape(Me.Form.Preview, Colors(PreviewColor), Me.NextShape)
        For Each xTile As Tile In PreviewShape.Tiles
            xTile.Left -= 2 * Tile.x
            xTile.Top += 1 * Tile.y
        Next

    End Sub
    Public Sub RemoveCompletedLines()
        Do
            Dim RemoveList As New List(Of Control)
            For TileTop As Integer = y * Tile.y To 0 Step -Tile.y
                RemoveList.Clear()
                For TileLeft As Integer = 0 To x * Tile.x Step Tile.x
                    For Each Control As Control In Me.Controls
                        If Control.Top = TileTop AndAlso Control.Left = TileLeft Then
                            RemoveList.Add(Control)
                        End If
                    Next
                Next
                If RemoveList.Count = x Then
                    For Each Tile As Tile In RemoveList
                        Me.Controls.Remove(Tile)
                    Next
                    For Each Control As Control In Me.Controls
                        If Control.Top < TileTop Then
                            Control.Top += Tile.y
                        End If
                    Next
                    Continue Do
                End If
            Next
            Exit Do
        Loop
    End Sub
End Class

Public Class Shape
    Public Style As Styles
    Public Enum Styles
        I
        T
        O
        S
        Z
        J
        L
    End Enum
    Public Direction As Directions
    Public Enum Directions
        North
        East
        South
        West
    End Enum
    Public Tiles As New List(Of Tile)
    Public Grid As Grid
    Public Sub New(Grid As Grid, Color As Color, Style As Styles)
        Me.Grid = Grid
        Me.Style = Style
        Select Case Me.Style

            Case Styles.I
                For Position As Integer = 3 To 6
                    Dim Tile As New Tile(Me, Color, Position * Tile.x, 0)
                Next

            Case Styles.T
                For PositionY As Integer = 0 To 1
                    For PositionX As Integer  = 3 To 5

                        If PositionY = 0 OrElse (PositionY = 1 AndAlso PositionX = 4) Then
                            Dim Tile As New Tile(Me, Color, PositionX * Tile.x, PositionY * Tile.y)
                        End If
                    Next
                Next

            Case Styles.O
                For PositionY As Integer  = 0 To 1
                    For PositionX As Integer  = 4 To 5
                        Dim Tile As New Tile(Me, Color, PositionX * Tile.x, PositionY * Tile.y)
                    Next
                Next

            Case Styles.S
                For PositionY As Integer  = 0 To 1
                    For PositionX As Integer  = 3 To 5
                        If (PositionY = 0 AndAlso PositionX > 3) OrElse (PositionY = 1 AndAlso PositionX < 5) Then
                            Dim Tile As New Tile(Me, Color, PositionX * Tile.x, PositionY * Tile.y)
                        End If
                    Next
                Next

            Case Styles.Z
                For PositionY As Integer = 0 To 1
                    For PositionX As Integer = 3 To 5
                        If (PositionY = 0 AndAlso PositionX < 5) OrElse (PositionY = 1 AndAlso PositionX > 3) Then
                            Dim Tile As New Tile(Me, Color, PositionX * Tile.x, PositionY * Tile.y)
                        End If
                    Next
                Next

            Case Styles.J
                For PositionY As Integer  = 0 To 2
                    For PositionX As Integer  = 4 To 5
                        If (PositionY < 3 AndAlso PositionX > 4) OrElse PositionY = 2 Then
                            Dim Tile As New Tile(Me, Color, PositionX * Tile.x, PositionY * Tile.y)
                        End If
                    Next
                Next

            Case Styles.L
                For PositionY As Integer  = 0 To 2
                    For PositionX As Integer  = 4 To 5
                        If (PositionY < 3 AndAlso PositionX < 5) OrElse PositionY = 2 Then
                            Dim Tile As New Tile(Me, Color, PositionX * Tile.x, PositionY * Tile.y)
                        End If
                    Next
                Next

        End Select
    End Sub
    Public Sub RotateAntiClockwise()
        Me.RotateClockwise()
        Me.RotateClockwise()
        Me.RotateClockwise()
    End Sub
    Public Sub RotateClockwise()
        Select Case Me.Style

            Case Styles.I
                Select Case Me.Direction
                    Case Directions.North
                        If Me.Tiles(0).CanMove(2, -2) AndAlso
                           Me.Tiles(1).CanMove(1, -1) AndAlso
                           Me.Tiles(3).CanMove(-1, 1) Then

                            Me.Direction = Directions.East
                            Me.Tiles(0).DoMove(2, -2)
                            Me.Tiles(1).DoMove(1, -1)
                            Me.Tiles(3).DoMove(-1, 1)

                        End If

                    Case Directions.East
                        If Me.Tiles(0).CanMove(-2, 2) AndAlso
                           Me.Tiles(1).CanMove(-1, 1) AndAlso
                           Me.Tiles(3).CanMove(1, -1) Then

                            Me.Direction = Directions.North
                            Me.Tiles(0).DoMove(-2, 2)
                            Me.Tiles(1).DoMove(-1, 1)
                            Me.Tiles(3).DoMove(1, -1)

                        End If
                End Select

            Case Styles.T
                Select Case Me.Direction
                    Case Directions.North
                        If Me.Tiles(2).CanMove(-1, -1) Then
                            Me.Direction = Directions.East
                            Me.Tiles(2).DoMove(-1, -1)
                        End If

                    Case Directions.East
                        If Me.Tiles(3).CanMove(1, -1) Then
                            Me.Direction = Directions.South
                            Me.Tiles(3).DoMove(1, -1)
                        End If

                    Case Directions.South
                        If Me.Tiles(0).CanMove(1, 1) Then
                            Me.Direction = Directions.West
                            Me.Tiles(0).DoMove(1, 1)
                        End If

                    Case Directions.West
                        If Me.Tiles(2).CanMove(1, 1) AndAlso
                           Me.Tiles(3).CanMove(-1, 1) AndAlso
                           Me.Tiles(0).CanMove(-1, -1) Then

                            Me.Direction = Directions.North
                            Me.Tiles(2).DoMove(1, 1)
                            Me.Tiles(3).DoMove(-1, 1)
                            Me.Tiles(0).DoMove(-1, -1)

                        End If

                End Select

            Case Styles.S
                Select Case Me.Direction
                    Case Directions.North
                        If Me.Tiles(1).CanMove(-2, 0) AndAlso
                           Me.Tiles(2).CanMove(0, -2) Then
                            Me.Direction = Directions.East
                            Me.Tiles(1).DoMove(-2, 0)
                            Me.Tiles(2).DoMove(0, -2)
                        End If

                    Case Directions.East
                        If Me.Tiles(1).CanMove(2, 0) AndAlso
                           Me.Tiles(2).CanMove(0, 2) Then
                            Me.Direction = Directions.North
                            Me.Tiles(1).DoMove(2, 0)
                            Me.Tiles(2).DoMove(0, 2)
                        End If

                End Select

            Case Styles.Z
                Select Case Me.Direction
                    Case Directions.North
                        If Me.Tiles(3).CanMove(-2, 0) AndAlso
                           Me.Tiles(2).CanMove(0, -2) Then
                            Me.Direction = Directions.East
                            Me.Tiles(3).DoMove(-2, 0)
                            Me.Tiles(2).DoMove(0, -2)
                        End If

                    Case Directions.East
                        If Me.Tiles(3).CanMove(2, 0) AndAlso
                           Me.Tiles(2).CanMove(0, 2) Then
                            Me.Direction = Directions.North
                            Me.Tiles(3).DoMove(2, 0)
                            Me.Tiles(2).DoMove(0, 2)
                        End If

                End Select

            Case Styles.J
                Select Case Me.Direction
                    Case Directions.North
                        If Me.Tiles(0).CanMove(1, 1) AndAlso
                           Me.Tiles(2).CanMove(0, -1) AndAlso
                           Me.Tiles(3).CanMove(-1, -2) Then
                            Me.Direction = Directions.East
                            Me.Tiles(0).DoMove(1, 1)
                            Me.Tiles(2).DoMove(0, -1)
                            Me.Tiles(3).DoMove(-1, -2)
                        End If

                    Case Directions.East
                        If Me.Tiles(0).CanMove(-1, 1) AndAlso
                           Me.Tiles(2).CanMove(1, -1) AndAlso
                           Me.Tiles(3).CanMove(2, 0) Then
                            Me.Direction = Directions.South
                            Me.Tiles(0).DoMove(-1, 1)
                            Me.Tiles(2).DoMove(1, -1)
                            Me.Tiles(3).DoMove(2, 0)
                        End If

                    Case Directions.South
                        If Me.Tiles(0).CanMove(-1, -1) AndAlso
                           Me.Tiles(2).CanMove(1, 1) AndAlso
                           Me.Tiles(3).CanMove(0, 2) Then
                            Me.Direction = Directions.West
                            Me.Tiles(0).DoMove(-1, -1)
                            Me.Tiles(2).DoMove(1, 1)
                            Me.Tiles(3).DoMove(0, 2)
                        End If

                    Case Directions.West
                        If Me.Tiles(0).CanMove(1, -1) AndAlso
                           Me.Tiles(2).CanMove(-2, 1) AndAlso
                           Me.Tiles(3).CanMove(-1, 0) Then
                            Me.Direction = Directions.North
                            Me.Tiles(0).DoMove(1, -1)
                            Me.Tiles(2).DoMove(-2, 1)
                            Me.Tiles(3).DoMove(-1, -0)
                        End If

                End Select

            Case Styles.L
                Select Case Me.Direction
                    Case Directions.North
                        If Me.Tiles(0).CanMove(1, 1) AndAlso
                           Me.Tiles(2).CanMove(-1, 0) AndAlso
                           Me.Tiles(3).CanMove(-2, -1) Then
                            Me.Direction = Directions.East
                            Me.Tiles(0).DoMove(1, 1)
                            Me.Tiles(2).DoMove(-1, 0)
                            Me.Tiles(3).DoMove(-2, -1)
                        End If

                    Case Directions.East
                        If Me.Tiles(0).CanMove(-1, 1) AndAlso
                           Me.Tiles(2).CanMove(0, -2) AndAlso
                           Me.Tiles(3).CanMove(1, -1) Then
                            Me.Direction = Directions.South
                            Me.Tiles(0).DoMove(-1, 1)
                            Me.Tiles(2).DoMove(0, -2)
                            Me.Tiles(3).DoMove(1, -1)
                        End If

                    Case Directions.South
                        If Me.Tiles(0).CanMove(-1, -1) AndAlso
                           Me.Tiles(2).CanMove(2, 0) AndAlso
                           Me.Tiles(3).CanMove(1, 1) Then
                            Me.Direction = Directions.West
                            Me.Tiles(0).DoMove(-1, -1)
                            Me.Tiles(2).DoMove(2, 0)
                            Me.Tiles(3).DoMove(1, 1)
                        End If

                    Case Directions.West
                        If Me.Tiles(0).CanMove(1, -1) AndAlso
                           Me.Tiles(2).CanMove(-1, 2) AndAlso
                           Me.Tiles(3).CanMove(0, 1) Then
                            Me.Direction = Directions.North
                            Me.Tiles(0).DoMove(1, -1)
                            Me.Tiles(2).DoMove(-1, 2)
                            Me.Tiles(3).DoMove(0, 1)
                        End If

                End Select

        End Select
    End Sub
    Public Sub MoveLeft()
        If Me.CanMove(-1, 0) Then
            Me.DoMove(-1, 0)
        End If
    End Sub
    Public Sub MoveRight()
        If Me.CanMove(1, 0) Then
            Me.DoMove(1, 0)
        End If
    End Sub
    Public Function CanMove(x As Integer, y As Integer) As Boolean
        For Each Tile As Tile In Me.Tiles
            If Not Tile.CanMove(x, y) Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Sub DoMove(x As Integer, y As Integer)
        For Each Tile As Tile In Me.Tiles
            Tile.DoMove(x, y)
        Next
    End Sub
    Public Sub Drop()
        Do While Me.CanMove(0, 1)
            Me.DoMove(0, 1)
        Loop
    End Sub
End Class

Public Class Tile
    Inherits Label
    Public Const x = 30
    Public Const y = 30
    Public Shape As Shape
    Public Sub New(Shape As Shape, Color As Color, Left As Integer, Top As Integer)
        Me.Left = Left
        Me.Top = Top
        Me.Width = x
        Me.Height = y
        Me.Shape = Shape
        Me.BackColor = Color
        Me.BorderStyle = BorderStyle.Fixed3D
        Me.Shape.Tiles.Add(Me)
        Me.Shape.Grid.Controls.Add(Me)
        Me.TextAlign = ContentAlignment.MiddleCenter
        'Me.Text = Me.Shape.Tiles.IndexOf(Me)
    End Sub
    Public Sub DoMove(x As Integer, y As Integer)
        Me.Location = New Point(Me.Left + x * Tile.x, Me.Top + y * Tile.y)
    End Sub
    Public Function CanMove(x As Integer, y As Integer) As Boolean
        Dim Location As New Point(Me.Left + x * Tile.x, Me.Top + y * Tile.y)

        'Check for Shape Contact
        For Each Tile As Tile In Me.Parent.Controls
            If Not Shape.Tiles.Contains(Tile) Then
                If Location = Tile.Location Then
                    Return False
                End If
            End If
        Next

        'Check for Grid Bounds
        If Location.X < 0 OrElse Location.X = Grid.x * Tile.x OrElse Location.Y = Grid.y * Tile.y Then
            Return False
        End If

        Return True
    End Function
End Class