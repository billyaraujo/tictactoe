
Public Class Form1

    Dim PlayMatrix(2, 2) As Char
    Dim PlayImage(2, 2) As PictureBox
    Dim LastPlay As Char
    Dim StopPlaying As Boolean
    Dim PlayAgainstCPU As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ReloadGame()

    End Sub

    Private Sub ReloadGame()

        PlayImage(0, 0) = PictureBox1
        PlayImage(0, 1) = PictureBox2
        PlayImage(0, 2) = PictureBox3
        PlayImage(1, 0) = PictureBox4
        PlayImage(1, 1) = PictureBox5
        PlayImage(1, 2) = PictureBox6
        PlayImage(2, 0) = PictureBox7
        PlayImage(2, 1) = PictureBox8
        PlayImage(2, 2) = PictureBox9

        LastPlay = "r"

        For i = 0 To 2
            For j = 0 To 2
                PlayMatrix(i, j) = "e"
                PlayImage(i, j).Image = Nothing
            Next
        Next

        StopPlaying = False
        PlayAgainstCPU = CheckBox1.Checked

    End Sub

    Private Sub DrawGame()

        ListBox1.Items.Insert(0, "Draw game.")
        My.Computer.Audio.Play(My.Resources.boring, AudioPlayMode.Background)
        StopPlaying = True

    End Sub

    Private Sub PlayerWins(Player As Char)

        If (Player = "b") Then
            ListBox1.Items.Insert(0, "Player blue wins.")
            If (PlayAgainstCPU) Then
                My.Computer.Audio.Play(My.Resources.woohoo, AudioPlayMode.Background)
            Else
                My.Computer.Audio.Play(My.Resources.woohoo, AudioPlayMode.Background)
            End If
            StopPlaying = True
        ElseIf (Player = "r") Then
            ListBox1.Items.Insert(0, "Player red wins.")
            If (PlayAgainstCPU) Then
                My.Computer.Audio.Play(My.Resources.doh, AudioPlayMode.Background)
            Else
                My.Computer.Audio.Play(My.Resources.woohoo, AudioPlayMode.Background)
            End If
            StopPlaying = True
        End If

    End Sub

    Private Sub CheckBoard()

        For i = 0 To 2
            If (PlayMatrix(i, 0) = "r") And (PlayMatrix(i, 1) = "r") And (PlayMatrix(i, 2) = "r") Then
                PlayerWins("r")
                Return
            End If
            If (PlayMatrix(i, 0) = "b") And (PlayMatrix(i, 1) = "b") And (PlayMatrix(i, 2) = "b") Then
                PlayerWins("b")
                Return
            End If
        Next

        For j = 0 To 2
            If (PlayMatrix(0, j) = "r") And (PlayMatrix(1, j) = "r") And (PlayMatrix(2, j) = "r") Then
                PlayerWins("r")
                Return
            End If
            If (PlayMatrix(0, j) = "b") And (PlayMatrix(1, j) = "b") And (PlayMatrix(2, j) = "b") Then
                PlayerWins("b")
                Return
            End If
        Next

        If (PlayMatrix(0, 0) = "r") And (PlayMatrix(1, 1) = "r") And (PlayMatrix(2, 2) = "r") Then
            PlayerWins("r")
            Return
        End If

        If (PlayMatrix(0, 2) = "r") And (PlayMatrix(1, 1) = "r") And (PlayMatrix(2, 0) = "r") Then
            PlayerWins("r")
            Return
        End If

        If (PlayMatrix(0, 0) = "b") And (PlayMatrix(1, 1) = "b") And (PlayMatrix(2, 2) = "b") Then
            PlayerWins("b")
            Return
        End If

        If (PlayMatrix(0, 2) = "b") And (PlayMatrix(1, 1) = "b") And (PlayMatrix(2, 0) = "b") Then
            PlayerWins("b")
            Return
        End If

        If (Not StopPlaying)

            Dim c As Integer

            c = 0

            For i = 0 To 2
                For j = 0 To 2
                    If PlayMatrix(i, j) <> "e" Then
                        c = c + 1
                    End If
                Next
            Next

            If (c = 9) Then
                DrawGame()
                Return
            End If
        End If

    End Sub

    Private Function CheckPlayerWin(Player As Char, ByRef ii As Integer, ByRef jj As Integer) As Boolean

        Dim p, e As Integer
        Dim j As Integer

        ' Check rows
        For i = 0 To 2

            p = 0
            e = -1

            For j = 0 To 2

                If (PlayMatrix(i, j) = Player) Then

                    p = p + 1

                ElseIf (PlayMatrix(i, j) = "e") Then

                    e = j

                    ii = i
                    jj = e

                End If
            Next

            If (p = 2 And e <> -1) Then
                Return True
            End If

        Next

        ' Check cols
        For j = 0 To 2

            p = 0
            e = -1

            For i = 0 To 2

                If (PlayMatrix(i, j) = Player) Then

                    p = p + 1

                ElseIf (PlayMatrix(i, j) = "e") Then

                    e = i

                    ii = e
                    jj = j

                End If
            Next

            If (p = 2 And e <> -1) Then
                Return True
            End If

        Next

        ' Check diagonals
        ' Check 1 diagonal
        p = 0
        e = -1

        For i = 0 To 2

            j = i

            If (PlayMatrix(i, j) = Player) Then

                p = p + 1

            ElseIf (PlayMatrix(i, j) = "e") Then

                e = j

                ii = i
                jj = e

            End If

        Next

        If (p = 2 And e <> -1) Then
            Return True
        End If

        ' Check 2 diagonal
        p = 0
        e = -1

        For i = 0 To 2

            j = 2 - i

            If (PlayMatrix(i, j) = Player) Then

                p = p + 1

            ElseIf (PlayMatrix(i, j) = "e") Then

                e = j

                ii = i
                jj = e

            End If

        Next

        If (p = 2 And e <> -1) Then
            Return True
        End If

        Return False

    End Function

    Private Function CheckCorner(ci As Integer, cj As Integer)

        Dim c As Integer

        c = 0

        If (PlayMatrix(ci, cj) = "e") Then

            For i = 0 To 2
                If PlayMatrix(i, cj) = "b" Then
                    c = c + 1
                End If
            Next

            For j = 0 To 2
                If PlayMatrix(ci, j) = "b" Then
                    c = c + 1
                End If
            Next

        End If

        Return c

    End Function

    Private Sub ComputerPlay()

        Dim i, j As Integer

        'First try to win
        If (CheckPlayerWin("r", i, j)) Then
            PlayPosition(i, j)
            Return
        End If

        'Second stop opponent
        If (CheckPlayerWin("b", i, j)) Then
            PlayPosition(i, j)
            Return
        End If

        ' Third check corners
        If (PlayMatrix(0, 0) = "b" And PlayMatrix(2, 2) = "b") Then
            If (PlayMatrix(0, 1) = "e") Then
                PlayPosition(0, 1)
                Return
            End If
            If (PlayMatrix(1, 2) = "e") Then
                PlayPosition(1, 2)
                Return
            End If
            If (PlayMatrix(1, 0) = "e") Then
                PlayPosition(1, 0)
                Return
            End If
            If (PlayMatrix(2, 1) = "e") Then
                PlayPosition(2, 1)
                Return
            End If
        End If

        If (PlayMatrix(2, 0) = "b" And PlayMatrix(0, 2) = "b") Then
            If (PlayMatrix(0, 1) = "e") Then
                PlayPosition(0, 1)
                Return
            End If
            If (PlayMatrix(1, 2) = "e") Then
                PlayPosition(1, 2)
                Return
            End If
            If (PlayMatrix(1, 0) = "e") Then
                PlayPosition(1, 0)
                Return
            End If
            If (PlayMatrix(2, 1) = "e") Then
                PlayPosition(2, 1)
                Return
            End If
        End If

        'Third play middle
        If (PlayMatrix(1, 1) = "e") Then

            PlayPosition(1, 1)
            Return

        End If

        'Fourth play corners
        Dim c1, c2, c3, c4 As Integer

        c1 = CheckCorner(0, 0)
        c2 = CheckCorner(0, 2)
        c3 = CheckCorner(2, 2)
        c4 = CheckCorner(2, 0)

        If (c1 >= c2 And c1 >= c3 And c1 >= c4) Then
            PlayPosition(0, 0)
            Return
        End If

        If (c2 >= c1 And c2 >= c3 And c2 >= c4) Then
            PlayPosition(0, 2)
            Return
        End If

        If (c3 >= c1 And c3 >= c2 And c3 >= c4) Then
            PlayPosition(2, 2)
            Return
        End If

        If (c4 >= c1 And c4 >= c2 And c4 >= c3) Then
            PlayPosition(2, 0)
            Return
        End If

        ' Fifth play middle edges
        If (PlayMatrix(0, 1) = "e") Then

            PlayPosition(0, 1)
            Return

        End If

        If (PlayMatrix(1, 2) = "e") Then

            PlayPosition(1, 2)
            Return

        End If

        If (PlayMatrix(2, 1) = "e") Then

            PlayPosition(2, 1)
            Return

        End If

        If (PlayMatrix(1, 0) = "e") Then

            PlayPosition(1, 0)
            Return

        End If

    End Sub

    Private Sub PlayPosition(i As Integer, j As Integer)

        If (PlayMatrix(i, j) = "e") And Not StopPlaying Then

            If (LastPlay = "r") Then
                PlayMatrix(i, j) = "b"
                PlayImage(i, j).Image = My.Resources.blue
                LastPlay = "b"
            ElseIf (LastPlay = "b") Then
                PlayMatrix(i, j) = "r"
                PlayImage(i, j).Image = My.Resources.red
                LastPlay = "r"
            End If

            CheckBoard()

            If (PlayAgainstCPU And LastPlay = "b") Then

                ComputerPlay()

            End If

        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        PlayPosition(0, 0)

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        PlayPosition(0, 1)

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

        PlayPosition(0, 2)

    End Sub


    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        PlayPosition(1, 0)

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click

        PlayPosition(1, 1)

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        PlayPosition(1, 2)

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click

        PlayPosition(2, 0)

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        PlayPosition(2, 1)

    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

        PlayPosition(2, 2)

    End Sub

    Private Sub PictureBox1_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox1.MouseEnter

        PlayImage(0, 0).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(0, 0)"

    End Sub

    Private Sub PictureBox1_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox1.MouseLeave

        PlayImage(0, 0).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox2.MouseEnter

        PlayImage(0, 1).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(0, 1)"

    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave

        PlayImage(0, 1).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox3_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox3.MouseEnter

        PlayImage(0, 2).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(0, 2)"

    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox3.MouseLeave

        PlayImage(0, 2).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox4_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox4.MouseEnter

        PlayImage(1, 0).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(1, 0)"

    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave

        PlayImage(1, 0).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox5_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox5.MouseEnter

        PlayImage(1, 1).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(1, 1)"

    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox5.MouseLeave

        PlayImage(1, 1).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox6_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox6.MouseEnter

        PlayImage(1, 2).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(1, 2)"

    End Sub

    Private Sub PictureBox6_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox6.MouseLeave

        PlayImage(1, 2).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox7_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox7.MouseEnter

        PlayImage(2, 0).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(2, 0)"

    End Sub

    Private Sub PictureBox7_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox7.MouseLeave

        PlayImage(2, 0).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox8_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox8.MouseEnter

        PlayImage(2, 1).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(2, 1)"

    End Sub

    Private Sub PictureBox8_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox8.MouseLeave

        PlayImage(2, 1).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub PictureBox9_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox9.MouseEnter

        PlayImage(2, 2).BackColor = Color.LightGray
        ToolStripStatusLabel1.Text = "(2, 2)"

    End Sub

    Private Sub PictureBox9_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox9.MouseLeave

        PlayImage(2, 2).BackColor = Control.DefaultBackColor
        ToolStripStatusLabel1.Text = ""

    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click

        ReloadGame()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        PlayAgainstCPU = CheckBox1.Checked

    End Sub

End Class
